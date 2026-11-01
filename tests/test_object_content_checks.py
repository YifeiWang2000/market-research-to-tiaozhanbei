"""Regression for clipped inline objects and correctly formatted but shifted cells."""
import tempfile,unittest,zipfile
from pathlib import Path
from test_check_delivery import c, budget, para, table, NS
WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
M='http://schemas.openxmlformats.org/officeDocument/2006/math'
def picture(ppr='',cy='1270000',anchor=False):
 kind='anchor' if anchor else 'inline'
 return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:drawing><wp:{kind}><wp:extent cx="1270000" cy="{cy}"/></wp:{kind}></w:drawing></w:r></w:p>'
def spacing(rule='exact',line='390'):return f'<w:spacing w:lineRule="{rule}" w:line="{line}"/>'
def style(name='Normal',props='',parent=None,default=False):
 return f'<w:style w:type="paragraph" w:styleId="{name}"'+(' w:default="1"' if default else '')+'>'+ (f'<w:basedOn w:val="{parent}"/>' if parent else '')+f'<w:pPr>{props}</w:pPr></w:style>'
class ObjectChecks(unittest.TestCase):
 def run_case(self,body,styles=None,reference=None):
  with tempfile.TemporaryDirectory() as tmp:
   p=Path(tmp)/'文档.docx'
   with zipfile.ZipFile(p,'w') as z:
    z.writestr('word/document.xml',f'<w:document xmlns:w="{NS}" xmlns:wp="{WP}" xmlns:m="{M}"><w:body>{para("说明",align="left")}{body}</w:body></w:document>')
    if styles is not None:z.writestr('word/styles.xml',f'<w:styles xmlns:w="{NS}">{styles}</w:styles>')
   return c.check(p,budget(),reference)
 def clipped(self,r):return any('裁切风险' in s for s in r['issues'])
 def test_direct_exact(self):self.assertTrue(self.clipped(self.run_case(picture(spacing()))))
 def test_inherit_default(self):self.assertTrue(self.clipped(self.run_case(picture(),style(props=spacing(),default=True))))
 def test_based_on_chain(self):
  styles=style('Base',spacing())+style('Normal',parent='Base',default=True)
  self.assertTrue(self.clipped(self.run_case(picture(),styles)))
 def test_doc_defaults(self):
  styles=f'<w:docDefaults><w:pPrDefault><w:pPr>{spacing()}</w:pPr></w:pPrDefault></w:docDefaults>'
  self.assertTrue(self.clipped(self.run_case(picture(),styles)))
 def test_explicit_auto_overrides_inherited(self):
  r=self.run_case(picture(spacing('auto','240')),style(props=spacing(),default=True))
  self.assertFalse(self.clipped(r));self.assertEqual(r['object_layout']['fixed_spacing_paragraphs'],[])
 def test_child_auto_overrides_parent(self):
  r=self.run_case(picture(),style('Base',spacing())+style('Normal',spacing('auto','240'),'Base',True))
  self.assertFalse(self.clipped(r))
 def test_at_least_does_not_clip(self):self.assertFalse(self.clipped(self.run_case(picture(spacing('atLeast')))))
 def test_short_image_still_warns(self):
  r=self.run_case(picture(spacing(),cy='12700'));self.assertFalse(self.clipped(r));self.assertTrue(any('可见范围' in s for s in r['warnings']))
 def test_math_exact_warns(self):
  r=self.run_case('<w:p><w:pPr>'+spacing()+'</w:pPr><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>')
  self.assertTrue(any('公式使用固定行高' in s for s in r['warnings']))
 def test_anchor_not_inline(self):self.assertFalse(self.clipped(self.run_case(picture(spacing(),anchor=True))))
 def test_unknown_style_warns(self):
  r=self.run_case(picture('<w:pStyle w:val="missing"/>'));self.assertTrue(any('样式无法解析' in s for s in r['warnings']))
 def test_cycle_terminates(self):
  r=self.run_case(picture(),style('Normal',spacing(),'B',True)+style('B',parent='Normal'));self.assertTrue(any('循环' in s for s in r['warnings']))
 def test_bad_height_warns(self):self.assertTrue(any('高度' in s for s in self.run_case(picture(spacing(),cy='nan'))['warnings']))
 def reference(self,rows,verified=True):return {'tables':[{'table_index':1,'source_location':'合成源页表1','source_verified':verified,'expected_rows':rows}]}
 def test_snapshot_not_certification(self):
  r=self.run_case(table());self.assertEqual(r['table_content']['snapshot'][0]['rows'],[['文字'],['文字'],['文字']]);self.assertEqual(r['table_content']['unverified_indices'],[1])
 def test_matching_source(self):
  r=self.run_case(table(),reference=self.reference([['文字'],['文字'],['文字']]))
  self.assertEqual(r['table_content']['source_verified_matching_indices'],[1]);self.assertFalse(r['issues'])
 def test_shifted_symbol_is_detected(self):
  r=self.run_case(table().replace('文字','X2',1),reference=self.reference([['X1'],['文字'],['文字']]))
  self.assertTrue(any('内容与源页' in s for s in r['issues']));self.assertEqual(r['table_content']['unverified_indices'],[1])
 def test_minus_and_decimal_are_significant(self):
  for actual in ['0.5','-5']:
   with self.subTest(actual=actual):
    r=self.run_case(table().replace('文字',actual,1),reference=self.reference([['-0.5'],['文字'],['文字']]))
    self.assertTrue(any('内容与源页' in s for s in r['issues']))
 def test_whitespace_only_ignored(self):
  r=self.run_case(table(),reference=self.reference([['文 字'],['文字'],['文字']]))
  self.assertEqual(r['table_content']['source_verified_matching_indices'],[1])
 def test_unverified_reference_does_not_certify(self):
  r=self.run_case(table(),reference=self.reference([['文字'],['文字'],['文字']],False))
  self.assertEqual(r['table_content']['source_verified_matching_indices'],[]);self.assertTrue(any('未声明源页' in s for s in r['warnings']))
 def test_invalid_reference(self):
  for ref in [[],{}, {'tables':[None]},self.reference(None),self.reference([[]]),{'tables':[{'table_index':True}]}]:
   with self.subTest(ref=ref):self.assertTrue(self.run_case(table(),reference=ref)['issues'])
 def test_partial_reference(self):
  r=self.run_case(table()+table(),reference=self.reference([['文字'],['文字'],['文字']]))
  self.assertEqual(r['table_content']['unverified_indices'],[2]);self.assertTrue(any('仅覆盖' in s for s in r['warnings']))
 def test_duplicate_reference(self):
  ref=self.reference([['文字'],['文字'],['文字']]);ref['tables']*=2
  self.assertTrue(any('索引重复' in s for s in self.run_case(table(),reference=ref)['issues']))
if __name__=='__main__':unittest.main(verbosity=2)
