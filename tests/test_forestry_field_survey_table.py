import importlib.util, pathlib, unittest, zipfile, xml.etree.ElementTree as E
ROOT=pathlib.Path(__file__).resolve().parents[1]
SKILL=ROOT/'skills/forestry-field-survey-table'
spec=importlib.util.spec_from_file_location('field_survey',SKILL/'scripts/forestry_field_survey_table.py');m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
class SurveyTests(unittest.TestCase):
 def test_mapping_sort_and_unknowns(self):
  h=['小班编','面积','LIN_BAN','XIAO_BAN','LIN_ZHONG','YOU_SHI_SZ','TU_RANG_LX','PINGJUN_SG','PO_DU','塔号']
  rows,w=m.prepare(h,[['02',.1,'0327','00004','242','410001','122',0,9,'T2'],['01',.2,'0327','00003','111','200000','122',None,4,'T1']])
  self.assertEqual(rows[0][6],'01');self.assertEqual(rows[0][13],'水源林');self.assertEqual(rows[0][16],'油松')
  self.assertEqual(rows[1][4:6],['0327','00004']);self.assertEqual(rows[1][13],'能源林');self.assertEqual(rows[1][16],'410001');self.assertEqual(rows[1][27],'122')
  self.assertEqual(rows[1][18],0);self.assertIsNone(rows[0][18]);self.assertEqual(rows[1][26],9)
  for i in (0,1,2,22,23,29,30,31):self.assertIsNone(rows[1][i])
 def test_unknown_forest_type_stays_code(self):
  rows,w=m.prepare(['小班编','面积','LIN_ZHONG'],[['1',.1,'999']])
  self.assertEqual(rows[0][13],'999')
  self.assertIn('Untranslated LIN_ZHONG: 999',w)
 def test_invalid_inputs(self):
  for h,rows in [(['小班编','TBMJ'],[['1',2]]),(['小班编','面积'],[['1',None]]),(['小班编','面积'],[['1',-1]]),(['小班编','面积'],[['1','NaN']]),(['小班编','面积'],[['1',1],['1',2]]),(['小班编','面积','面积'],[['1',1,2]]),(['小班编','面积'],[])]:
   with self.subTest(h=h,rows=rows):
    with self.assertRaises(ValueError):m.prepare(h,rows)
 def test_explicit_fields_and_blank_rows(self):
  rows,_=m.prepare(['ID','HA','市'],[[None,None,None],['A',0,'已有市']],id_field='ID',area_field='HA',county='已确认县')
  self.assertEqual(rows[0][0:3],['已有市','已确认县',None]);self.assertEqual(rows[0][7],0)
 def test_default_output(self):
  self.assertEqual(m.default_output(pathlib.Path('/tmp/调查_初表.xls')).name,'调查_外调.xlsx')
 def test_template_classification_and_reference_stock(self):
  h=['小班编','面积','DLMC','LIN_ZHONG','SEN_LIN_LB','SHI_QUAN_D','MEI_GQ_XJ']
  rows,w=m.prepare(h,[['1',.1,'灌木林地','242','012','20',50],['2',.2,'乔木林地','112','012','10',80]])
  self.assertEqual(rows[0][8:10],['一般灌木林地','能源林林地'])
  self.assertEqual(rows[1][9],'防护林林地');self.assertEqual(rows[0][11],'一般公益林')
  self.assertEqual(rows[0][21],50)
  for r in rows:self.assertIsNone(r[22])
  self.assertFalse(any('SHI_QUAN_D' in x or '细分' in x or '未匹配林地类型' in x for x in w))
 def test_names_nursery_and_unknown_types(self):
  cases=[('一般灌木林地','经济林','经济林林地'),('特殊灌木林地','用材林','用材林林地'),('国家特别规定灌木林地','能源林','能源林林地'),('乔木林地','111','防护林林地'),('乔木林地','防护林','防护林林地'),('乔木林地','233','用材林林地'),('苗圃地',None,'苗圃地'),('其他林地','999',None)]
  for land,forest,expected in cases:
   with self.subTest(land=land,forest=forest):
    rows,w=m.prepare(['小班编','面积','DLMC','LIN_ZHONG'],[['1',.1,land,forest]])
    self.assertEqual(rows[0][9],expected)
    if '灌木' in land:self.assertEqual(rows[0][8],'一般灌木林地')
    if expected is None:self.assertEqual(rows[0][13],'999');self.assertTrue(any('未匹配林地类型' in x for x in w))
 def test_workflow_observations_are_independent(self):
  headers=['小班编','面积','SHI_QUAN_D','LIN_ZHONG','PINGJUN_SG','YU_BI_DU']
  raw=[['02',.1,'20','242',5,.15],['01',.2,'10','111',8,.7]]
  rows,_=m.prepare(headers,raw)
  sheets=m.workflow_rows(['_id']+headers[1:],raw,rows,'临时')
  self.assertEqual([s['name'] for s in sheets],['原始因子表','外业调查表'])
  original,field=[s['rows'] for s in sheets]
  self.assertEqual([len(original[0]),len(field[0])],[35,37])
  self.assertEqual(original[0][12],'国家公益林');self.assertEqual(original[1][12],'地方公益林')
  self.assertEqual(original[1][22],.15);self.assertEqual(original[1][33],'临时')
  for row in field:
   for i in [3,13,17,19,21,25,27,29,31,33,35]:self.assertIsNone(row[i])
  field[0][3]='疏林地';self.assertIsNone(original[0][8])
 def test_template_clean_and_styled(self):
  ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
  with zipfile.ZipFile(SKILL/'assets/A.1.xlsx') as z:
   self.assertNotIn(b'absPath',z.read('xl/workbook.xml'))
   wb=E.fromstring(z.read('xl/workbook.xml'));self.assertEqual([x.get('name') for x in wb.find('m:sheets',ns)],['原始因子表','外业调查表'])
   self.assertFalse(any('externalLink' in n or 'vbaProject' in n for n in z.namelist()))
   for i in (1,2):
    sh=E.fromstring(z.read(f'xl/worksheets/sheet{i}.xml'))
    self.assertIsNotNone(sh.find('m:pageSetup',ns))
    for cell in sh.findall('.//m:sheetData/m:row/m:c',ns):
     row=int(''.join(c for c in cell.get('r') if c.isdigit()))
     if row>=4:self.assertIsNone(cell.find('m:v',ns));self.assertIsNone(cell.find('m:f',ns))
   strings=E.fromstring(z.read('xl/sharedStrings.xml'));texts=[''.join(x.itertext()) for x in strings]
   self.assertNotIn('旧编号',texts)
   self.assertFalse(any('三家' in x or '承德' in x for x in texts))
if __name__=='__main__':unittest.main()
