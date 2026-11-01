"""Conservative delivery checks; never certifies research or rendered layout."""
import argparse
import hashlib
import json
import math
import re
import zipfile
import xml.etree.ElementTree as E
from datetime import datetime, timezone
from pathlib import Path

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def txt(node):
    return ''.join(x.text or '' for x in node.iter(W + 't'))
def find(node, path):
    return node.find('/'.join(W + x for x in path.split('/')))
def number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value) and value >= 0


WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'

def object_layout(root, styles, issues, warnings):
    """Resolve paragraph spacing only; never infer rendered visibility from OOXML."""
    catalog = {} if styles is None else {x.get(W+'styleId'): x for x in styles.findall(W+'style')}
    default_id = next((k for k, x in catalog.items() if x.get(W+'type') == 'paragraph'
                       and x.get(W+'default') in ('1', 'true', 'on')), None)
    base = None if styles is None else find(styles, 'docDefaults/pPrDefault/pPr/spacing')
    stats = {'inline_images': 0, 'math_paragraphs': 0, 'fixed_spacing_paragraphs': []}
    def spacing(para, label):
        attrs = {} if base is None else dict(base.attrib)
        st = find(para, 'pPr/pStyle')
        sid = st.get(W+'val') if st is not None else default_id
        chain, seen = [], set()
        while sid:
            if sid in seen:
                warnings.append(label+'样式继承存在循环，行距待人工核查'); break
            seen.add(sid)
            node = catalog.get(sid)
            if node is None:
                warnings.append(label+'样式无法解析，行距待人工核查'); break
            chain.append(node)
            parent = node.find(W+'basedOn')
            sid = parent.get(W+'val') if parent is not None else None
        for node in reversed(chain):
            value = find(node, 'pPr/spacing')
            if value is not None: attrs.update(value.attrib)
        value = find(para, 'pPr/spacing')
        if value is not None: attrs.update(value.attrib)
        return attrs
    for pi, para in enumerate(root.iter(W+'p'), 1):
        inline = list(para.iter(WP+'inline'))
        maths = list(para.iter(M+'oMath'))
        legacy = list(para.iter(W+'pict'))
        if not inline and not maths and not legacy: continue
        label = '第%d段对象：' % pi
        stats['inline_images'] += len(inline)
        stats['math_paragraphs'] += bool(maths)
        attrs = spacing(para, label)
        if attrs.get(W+'lineRule') != 'exact': continue
        stats['fixed_spacing_paragraphs'].append(pi)
        try:
            line_pt = float(attrs.get(W+'line', '')) / 20
            if not math.isfinite(line_pt) or line_pt <= 0: raise ValueError()
        except ValueError:
            warnings.append(label+'固定行高数值无效或缺失，需核查显示'); continue
        for obj in inline:
            extent = obj.find(WP+'extent')
            try:
                height_pt = float(extent.get('cy', '')) / 12700
                if not math.isfinite(height_pt) or height_pt <= 0: raise ValueError()
            except (ValueError, AttributeError):
                warnings.append(label+'图片高度缺失或无效，需渲染核查'); continue
            if height_pt > line_pt:
                issues.append(label+'内嵌图片高度%.2f磅超过固定行高%.2f磅，存在裁切风险' % (height_pt,line_pt))
            else:
                warnings.append(label+'内嵌图片使用固定行高，仍需核查基线及可见范围')
        if maths: warnings.append(label+'公式使用固定行高，分式/上下标等须渲染核查')
        if legacy: warnings.append(label+'旧式图形使用固定行高，无法自动确认浮动方式及可见范围')
    return stats

def table_content(root, reference, issues, warnings):
    """Expose physical XML cells; compare only against independently checked source rows."""
    tables = list(root.iter(W+'tbl'))
    snapshot = []
    for index, table in enumerate(tables, 1):
        rows, merges = [], []
        for ri, row in enumerate(table.findall(W+'tr'), 1):
            cells = row.findall(W+'tc'); rows.append([txt(x) for x in cells])
            for ci, cell in enumerate(cells, 1):
                span, merge = find(cell, 'tcPr/gridSpan'), find(cell, 'tcPr/vMerge')
                if span is not None or merge is not None:
                    merges.append({'row':ri,'cell':ci,'grid_span':span.get(W+'val') if span is not None else None,
                                   'vertical_merge':merge.get(W+'val','continue') if merge is not None else None})
        snapshot.append({'table_index':index,'rows':rows,'merges':merges})
    compared, verified = set(), set()
    if reference is not None:
        entries = reference.get('tables') if isinstance(reference, dict) else None
        if not isinstance(entries, list) or not entries:
            issues.append('表格内容基准tables必须为非空列表'); entries=[]
        seen=set()
        for item in entries:
            if not isinstance(item, dict): issues.append('表格内容基准条目必须为对象'); continue
            index=item.get('table_index'); expected=item.get('expected_rows')
            if type(index) is not int or not 1 <= index <= len(tables):
                issues.append('表格内容基准索引无效'); continue
            if index in seen: issues.append('表格内容基准索引重复：%d' % index); continue
            seen.add(index)
            if not isinstance(item.get('source_location'),str) or not item['source_location'].strip():
                issues.append('表对象%d内容基准缺源页位置' % index); continue
            if not isinstance(expected,list) or not expected or any(not isinstance(row,list) or not row or any(not isinstance(c,str) for c in row) for row in expected):
                issues.append('表对象%d内容基准必须为非空二维字符串数组' % index); continue
            # Only whitespace is ignored. Minus signs, decimals, labels and units remain significant.
            normalize=lambda rows:[[re.sub(r'\s','',c) for c in row] for row in rows]
            matches=normalize(expected)==normalize(snapshot[index-1]['rows'])
            compared.add(index)
            if not matches: issues.append('表对象%d内容与源页核对基准不一致，检查行列/符号/数值/单位' % index)
            if item.get('source_verified') is not True:
                warnings.append('表对象%d基准未声明源页核对，不能据匹配认定内容正确' % index)
            elif matches: verified.add(index)
        if len(compared)<len(tables): warnings.append('表格内容基准仅覆盖%d/%d个表对象，其余待核' % (len(compared),len(tables)))
    return {'snapshot':snapshot,'compared_indices':sorted(compared),'source_verified_matching_indices':sorted(verified),
            'unverified_indices':[n for n in range(1,len(tables)+1) if n not in verified],
            'scope':'物理单元格文字比较；合并关系、图像表格、源页真实性与解释仍需人工核对'}

def check(docx, budget, table_reference=None):
    issues, warnings = [], []
    manual = ['图内文字颜色', '继承样式及最终渲染', '实际页数与同口径字数来源',
              '逐章分析及图表保留', '表格内容与源页逐项核对', '重要主张全文及图内一致性', '引用准确性与学术表达', '学生自主撰写及提交资格']
    result = {'file': str(docx), 'checked_at': datetime.now(timezone.utc).isoformat(),
              'sha256': None, 'caption_counts': {'图': 0, '表': 0}}
    try:
        raw = Path(docx).read_bytes()
        result['sha256'] = hashlib.sha256(raw).hexdigest()
        with zipfile.ZipFile(docx) as z:
            root = E.fromstring(z.read('word/document.xml'))
            styles = None
            if 'word/styles.xml' in z.namelist():
                try: styles = E.fromstring(z.read('word/styles.xml'))
                except E.ParseError: warnings.append('样式文件无法解析，图片及公式行距须人工核查')
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, E.ParseError) as exc:
        root = None
        issues.append('DOCX无法读取：' + type(exc).__name__)
    if root is not None:
        result['object_layout'] = object_layout(root, styles, issues, warnings)
        result['table_content'] = table_content(root, table_reference, issues, warnings)
        if not txt(root).strip():
            issues.append('文档无可识别正文文字，需确认是否为空或仅含图像')
        caps = {'图': [], '表': []}
        for para in root.iter(W + 'p'):
            text = txt(para).strip()
            m = re.match(r'^([图表])\s*(\d+)(?:\s|[：:、.])', text)
            if not m:
                continue
            style = find(para, 'pPr/pStyle')
            is_caption = style is not None and style.get(W + 'val', '').lower() in ('caption', '题注')
            if not is_caption and (text.endswith(('。', '！', '？', '；')) or re.match(r'^[图表]\s*\d+\s*(显示|表明|说明|所示|中的|给出)', text)):
                warnings.append('疑似正文引用，未计作题注，需人工确认：' + text[:60])
                continue
            caps[m[1]].append(int(m[2]))
            jc = find(para, 'pPr/jc')
            if jc is None:
                warnings.append('题注对齐未显式设置，需核查继承：' + text[:60])
            elif jc.get(W + 'val') != 'center':
                issues.append('题注未居中：' + text[:60])
            for run in para.iter(W + 'r'):
                if not txt(run):
                    continue
                color = find(run, 'rPr/color')
                if color is None or color.get(W + 'val') == 'auto':
                    warnings.append('题注字色需核查继承：' + text[:60])
                elif color.get(W + 'val', '').upper() != '000000' or color.get(W + 'themeColor'):
                    issues.append('题注字色非显式纯黑或受主题覆盖：' + text[:60])
        for kind, nums in caps.items():
            result['caption_counts'][kind] = len(nums)
            if nums and nums != list(range(1, len(nums) + 1)):
                issues.append(kind + '编号非连续递增；分章编号需人工判定')
        for index, table in enumerate(root.iter(W + 'tbl'), 1):
            rows = table.findall(W + 'tr')
            if '研究思路框架图' in txt(table) and len(rows) == 1:
                manual.append('表对象%d疑似框架占位，需核实未被当成内容表跳过' % index)
                continue
            style = find(table, 'tblPr/tblStyle')
            if style is not None:
                warnings.append('表对象%d使用表格样式，需渲染核查继承底纹/边框' % index)
            for shading in table.iter(W + 'shd'):
                fill = shading.get(W + 'fill', 'auto').upper()
                if fill not in ('AUTO', 'FFFFFF') or shading.get(W + 'themeFill'):
                    issues.append('表对象%d存在底纹或主题底纹' % index)
            # Word repeats a leading contiguous header group, not arbitrary rows.
            header_end = 0
            leading = True
            for ri, row in enumerate(rows):
                flag = find(row, 'trPr/tblHeader')
                enabled = flag is not None and flag.get(W + 'val', 'true') not in ('0', 'false', 'off')
                if leading and enabled:
                    header_end = ri
                else:
                    leading = False
                    if enabled:
                        warnings.append('表对象%d表头标记不连续，需人工确认' % index)
            if rows and header_end == len(rows) - 1:
                warnings.append('表对象%d没有可识别的数据行，需人工确认' % index)
            for ri, row in enumerate(rows):
                for cell in row.findall(W + 'tc'):
                    borders = find(cell, 'tcPr/tcBorders')
                    if borders is None:
                        warnings.append('表对象%d单元格边框依赖继承，需人工确认' % index)
                    else:
                        expected = {'top': '12' if ri == 0 else None,
                                    'bottom': '12' if ri == len(rows)-1 else '4' if ri == header_end else None,
                                    'left': None, 'right': None, 'start': None, 'end': None,
                                    'insideH': None, 'insideV': None}
                        for edge, size in expected.items():
                            edge_node = borders.find(W + edge)
                            val = edge_node.get(W + 'val') if edge_node is not None else None
                            label = '表对象%d第%d行%s' % (index, ri+1, edge)
                            if size:
                                if edge_node is None:
                                    warnings.append(label + '线未显式设置，需核查继承')
                                elif val != 'single' or edge_node.get(W + 'sz') != size or edge_node.get(W + 'color', '').upper() != '000000' or edge_node.get(W + 'themeColor'):
                                    issues.append(label + '线不符合默认三线表值')
                            elif val not in (None, 'none', 'nil'):
                                issues.append(label + '存在额外线')
                    for run in cell.iter(W + 'r'):
                        if txt(run):
                            color = find(run, 'rPr/color')
                            if color is None or color.get(W + 'val') == 'auto':
                                warnings.append('表对象%d文字颜色需核查继承' % index)
                            elif color.get(W + 'val', '').upper() != '000000' or color.get(W + 'themeColor'):
                                issues.append('表对象%d文字非显式纯黑或受主题覆盖' % index)
    if not isinstance(budget, dict):
        issues.append('预算顶层必须为对象')
        budget = {}
    units = budget.get('units')
    if not isinstance(units, list) or not units:
        issues.append('预算units必须为非空章节列表')
        units = []
    if budget.get('final_file_hash') and budget['final_file_hash'] != result['sha256']:
        issues.append('预算登记的最终文件哈希与受检文件不一致')
    for i, unit in enumerate(units):
        if not isinstance(unit, dict):
            issues.append('预算第%d项必须为对象' % (i+1))
            continue
        name = unit.get('name')
        if not isinstance(name, str) or not name.strip():
            issues.append('预算第%d项缺有效名称' % (i+1))
            name = '条目%d' % (i+1)
        for metric in ('text_units', 'pages'):
            spec = unit.get(metric)
            label = name + '/' + metric
            if not isinstance(spec, dict):
                issues.append(label + '必须为指标对象')
                continue
            for key in ('source', 'actual', 'min', 'max'):
                value = spec.get(key)
                if value is not None and not number(value):
                    issues.append(label + '/' + key + '须为有限非负数，不能为布尔值')
            if any(spec.get(k) is not None and not number(spec[k]) for k in ('source', 'actual', 'min', 'max')):
                continue
            applicability = spec.get('applicable', True)
            if not isinstance(applicability, bool):
                issues.append(label + '/applicable须为布尔值')
                continue
            if not applicability:
                if not isinstance(spec.get('reason'), str) or not spec['reason'].strip():
                    issues.append(label + '不适用须填写原因')
                else:
                    warnings.append(label + '声明不适用，需核对原因：' + spec['reason'])
                continue
            mode = spec.get('enforcement', 'review')
            if mode not in ('review', 'strict'):
                issues.append(label + 'enforcement须为review或strict')
                continue
            actual, low, high = (spec.get(k) for k in ('actual', 'min', 'max'))
            if actual is None or (low is None and high is None):
                warnings.append(label + '预算或实测缺失，不能完成该项验收')
                continue
            if low is not None and high is not None and low > high:
                issues.append(label + '目标上下限倒置')
                continue
            comparable = metric != 'pages' or unit.get('comparable_layout') is True
            if not comparable:
                warnings.append(label + '版式不可比或未确认，不据页数判通过/不通过')
            outside = (low is not None and actual < low) or (high is not None and actual > high)
            if outside:
                message = label + '超出事前目标；须核查内容、图表及偏差原因'
                (issues if mode == 'strict' and comparable else warnings).append(message)
    result.update(issues=list(dict.fromkeys(issues)), warnings=list(dict.fromkeys(warnings)),
                  mechanical_status='需修正或核查' if issues or warnings else '所检机械项未发现问题',
                  manual_review_required=list(dict.fromkeys(manual)), overall_status='本工具不判定整篇验收通过')
    return result

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    for key in ('docx', 'budget', 'output'):
        parser.add_argument('--' + key, required=True)
    parser.add_argument('--table-reference', help='可选：经源页核对的物理单元格内容基准JSON')
    args = parser.parse_args()
    try:
        budget = json.loads(Path(args.budget).read_text(encoding='utf-8'))
        error = None
    except (OSError, ValueError) as exc:
        budget, error = {}, '预算文件无法解析：' + type(exc).__name__
    table_reference, table_error = None, None
    if args.table_reference:
        try: table_reference = json.loads(Path(args.table_reference).read_text(encoding='utf-8'))
        except (OSError, ValueError) as exc: table_error = '表格内容基准无法解析：' + type(exc).__name__
    report = check(args.docx, budget, table_reference)
    if table_error:
        report['issues'].insert(0, table_error)
        report['mechanical_status'] = '需修正或核查'
    if error:
        report['issues'].insert(0, error)
        report['mechanical_status'] = '需修正或核查'
    Path(args.output).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(report['mechanical_status'], '问题', len(report['issues']), '待核', len(report['warnings']))
