"""Minimal OOXML fixtures test checker behavior, not Word rendering."""
import copy, importlib.util, json, tempfile, unittest, zipfile
from pathlib import Path
spec=importlib.util.spec_from_file_location('checker',Path(__file__).resolve().parents[1]/'scripts/check_delivery.py')
c=importlib.util.module_from_spec(spec);spec.loader.exec_module(c)
NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def para(text='图1 测试',align='center',color='000000'):
 return f'<w:p><w:pPr><w:jc w:val="{align}"/></w:pPr><w:r><w:rPr><w:color w:val="{color}"/></w:rPr><w:t>{text}</w:t></w:r></w:p>'
def table(extra=False,shade=False):
 s='<w:tbl>'
 for i in range(3):
  s+='<w:tr>'+('<w:trPr><w:tblHeader/></w:trPr>' if i<2 else '')+'<w:tc><w:tcPr><w:tcBorders>'
  for edge in ('top','bottom','left','right','insideH','insideV'):
   size='12' if (edge=='top' and i==0) or (edge=='bottom' and i==2) else '4' if edge=='bottom' and (i==1 or extra and i==0) else None
   s+=f'<w:{edge} w:val="single" w:sz="{size}" w:color="000000"/>' if size else f'<w:{edge} w:val="nil"/>'
  s+='</w:tcBorders>'+('<w:shd w:fill="0000FF"/>' if shade else '')+'</w:tcPr>'+para('文字')+'</w:tc></w:tr>'
 return s+'</w:tbl>'
def budget():
 return {'units':[{'name':'正文','comparable_layout':True,'text_units':{'min':1,'max':10,'actual':5},'pages':{'min':1,'max':2,'actual':1}}]}
class Checks(unittest.TestCase):
 def run_case(self,b=None,body=None):
  with tempfile.TemporaryDirectory() as d:
   path=Path(d)/'输入 文件.docx'
   with zipfile.ZipFile(path,'w') as z:z.writestr('word/document.xml',f'<w:document xmlns:w="{NS}"><w:body>{para() if body is None else body}</w:body></w:document>')
   return c.check(path,budget() if b is None else b)
 def test_normal(self):self.assertEqual(self.run_case()['issues'],[])
 def test_schema(self):
  for b in [[],{'units':None},{'units':'x'},{'units':[4]},{'units':[{}]},{'units':[{'name':'a','pages':None,'text_units':[]}]}]:
   with self.subTest(b=b):self.assertTrue(self.run_case(b)['issues'])
 def test_bad_numbers(self):
  for value in [-1,True,float('inf'),float('nan'),'1']:
   b=budget();b['units'][0]['pages']['actual']=value
   with self.subTest(value=value):self.assertTrue(self.run_case(b)['issues'])
 def test_reversed(self):
  b=budget();b['units'][0]['pages'].update(min=3,max=1);self.assertTrue(self.run_case(b)['issues'])
 def test_single_bounds(self):
  for key in ['min','max']:
   b=budget();b['units'][0]['pages'][key]=None;self.assertFalse(self.run_case(b)['issues'])
 def test_not_applicable(self):
  b=budget();b['units'][0]['pages']={'applicable':False,'reason':'Markdown无固定分页'};r=self.run_case(b);self.assertFalse(r['issues']);self.assertTrue(r['warnings'])
 def test_not_applicable_requires_reason(self):
  b=budget();b['units'][0]['pages']={'applicable':False};self.assertTrue(self.run_case(b)['issues'])
 def test_applicable_type(self):
  b=budget();b['units'][0]['pages']['applicable']='false';self.assertTrue(self.run_case(b)['issues'])
 def test_review_deviation(self):
  b=budget();b['units'][0]['pages']['actual']=3;r=self.run_case(b);self.assertFalse(r['issues']);self.assertTrue(r['warnings'])
 def test_strict_deviation(self):
  b=budget();b['units'][0]['pages'].update(actual=3,enforcement='strict');self.assertTrue(self.run_case(b)['issues'])
 def test_noncomparable(self):
  b=budget();b['units'][0].update(comparable_layout=False);b['units'][0]['pages'].update(actual=3,enforcement='strict');r=self.run_case(b);self.assertFalse(r['issues']);self.assertTrue(r['warnings'])
 def test_no_actual(self):
  b=budget();b['units'][0]['pages']['actual']=None;self.assertTrue(self.run_case(b)['warnings'])
 def test_empty(self):self.assertTrue(self.run_case(body='')['issues'])
 def test_red_caption(self):self.assertTrue(self.run_case(body=para(color='FF0000'))['issues'])
 def test_left_caption(self):self.assertTrue(self.run_case(body=para(align='left'))['issues'])
 def test_gap(self):self.assertTrue(self.run_case(body=para('图2 测试'))['issues'])
 def test_prose(self):
  r=self.run_case(body=para('图1 显示了调查范围。',align='left'));self.assertFalse(r['issues']);self.assertEqual(r['caption_counts']['图'],0);self.assertTrue(r['warnings'])
 def test_two_header_rows(self):self.assertFalse(self.run_case(body=table())['issues'])
 def test_extra_header_line(self):self.assertTrue(self.run_case(body=table(extra=True))['issues'])
 def test_shading(self):self.assertTrue(self.run_case(body=table(shade=True))['issues'])
 def test_stale_hash(self):
  b=budget();b['final_file_hash']='bad';self.assertTrue(self.run_case(b)['issues'])
 def test_missing_doc(self):self.assertTrue(c.check('/missing.docx',budget())['issues'])
if __name__=='__main__':unittest.main(verbosity=2)
