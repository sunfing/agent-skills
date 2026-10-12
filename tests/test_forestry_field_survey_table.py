import importlib.util, pathlib, unittest, zipfile, xml.etree.ElementTree as E
ROOT=pathlib.Path(__file__).resolve().parents[1]
SKILL=ROOT/'skills/forestry-field-survey-table'
spec=importlib.util.spec_from_file_location('field_survey',SKILL/'scripts/forestry_field_survey_table.py');m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
class SurveyTests(unittest.TestCase):
 def test_mapping_sort_and_unknowns(self):
  h=['小班编','面积','LIN_BAN','XIAO_BAN','LIN_ZHONG','YOU_SHI_SZ','TU_RANG_LX','PINGJUN_SG','PO_DU','塔号']
  rows,w=m.prepare(h,[['02',.1,'0327','00004','242','410001','122',0,9,'T2'],['01',.2,'0327','00003','111','200000','122',None,4,'T1']])
  self.assertEqual(rows[0][6],'01');self.assertEqual(rows[0][13],'水源林');self.assertEqual(rows[0][16],'油松')
  self.assertEqual(rows[1][4:6],['0327','00004']);self.assertEqual(rows[1][13],'242');self.assertEqual(rows[1][16],'410001');self.assertEqual(rows[1][27],'122')
  self.assertEqual(rows[1][18],0);self.assertIsNone(rows[0][18]);self.assertEqual(rows[1][26],9)
  for i in (0,1,2,9,22,23,29,30,31):self.assertIsNone(rows[1][i])
 def test_invalid_inputs(self):
  for h,rows in [(['小班编','TBMJ'],[['1',2]]),(['小班编','面积'],[['1',None]]),(['小班编','面积'],[['1',-1]]),(['小班编','面积'],[['1','NaN']]),(['小班编','面积'],[['1',1],['1',2]]),(['小班编','面积','面积'],[['1',1,2]]),(['小班编','面积'],[])]:
   with self.subTest(h=h,rows=rows):
    with self.assertRaises(ValueError):m.prepare(h,rows)
 def test_explicit_fields_and_blank_rows(self):
  rows,_=m.prepare(['ID','HA','市'],[[None,None,None],['A',0,'已有市']],id_field='ID',area_field='HA',county='已确认县')
  self.assertEqual(rows[0][0:3],['已有市','已确认县',None]);self.assertEqual(rows[0][7],0)
 def test_default_output(self):
  self.assertEqual(m.default_output(pathlib.Path('/tmp/调查_初表.xls')).name,'调查_外调.xlsx')
 def test_template_clean_and_styled(self):
  ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
  with zipfile.ZipFile(SKILL/'assets/A.1.xlsx') as z:
   self.assertNotIn(b'absPath',z.read('xl/workbook.xml'))
   wb=E.fromstring(z.read('xl/workbook.xml'));self.assertEqual([x.get('name') for x in wb.find('m:sheets',ns)],['A.1'])
   self.assertFalse(any('externalLink' in n or 'vbaProject' in n for n in z.namelist()))
   sh=E.fromstring(z.read('xl/worksheets/sheet1.xml'))
   for cell in sh.findall('.//m:sheetData/m:row/m:c',ns):
    row=int(''.join(c for c in cell.get('r') if c.isdigit()))
    if row>=4:self.assertIsNone(cell.find('m:v',ns));self.assertIsNone(cell.find('m:f',ns))
   self.assertEqual({c.get('ref') for c in sh.findall('m:mergeCells/m:mergeCell',ns)},{'A1:AG1','A2:C2'})
   strings=E.fromstring(z.read('xl/sharedStrings.xml'))
   texts=[''.join(x.itertext()) for x in strings]
   self.assertLessEqual(len(texts),35) # title, table label, 33 column labels only
   self.assertIn('使用林地小班因子调查表',texts)
if __name__=='__main__':unittest.main()
