import importlib.util, pathlib, unittest, zipfile, xml.etree.ElementTree as E, tempfile, shutil
ROOT=pathlib.Path(__file__).resolve().parents[1]
SKILL=ROOT/'skills/forestry-field-survey-table'
spec=importlib.util.spec_from_file_location('field_survey',SKILL/'scripts/forestry_field_survey_table.py');m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
class SurveyTests(unittest.TestCase):
 def test_previously_missing_common_codes(self):
  h=['序号','面积','BH_DJ','LING_ZU','LIN_ZHONG','SEN_LIN_LB','YOU_SHI_SZ','LD_QS','QI_YUAN']
  rows,w=m.prepare(h,[['01',.0126,'2','3','251','011','706000','21','12']])
  r=rows[0]
  self.assertEqual([r[i] for i in [9,10,11,12,13,14,16,17]],['经济林林地','集体','重点公益林','Ⅱ','果树林','天然','杏','近熟林'])
  self.assertFalse(any(x.startswith('Untranslated') for x in w))
 def test_all_forest_groups_and_names(self):
  groups=[(range(111,118),'防护林林地'),(range(121,128),'特用林林地'),(range(231,234),'用材林林地'),([240],'薪炭林林地'),([242],'能源林林地'),(range(251,256),'经济林林地')]
  for codes,typ in groups:
   for code in codes:
    with self.subTest(code=code):
     for v in [str(code),m.CODE_FIELDS['LIN_ZHONG']['codes'][str(code)]]:
      rows,w=m.prepare(['序号','面积','LIN_ZHONG'],[['1',.1,v]])
      self.assertEqual(rows[0][9],typ)
      self.assertFalse(any('Untranslated LIN_ZHONG' in x for x in w))
  self.assertEqual(m.CODE_FIELDS['LIN_ZHONG']['entry_sources']['242'],'confirmed-template')
 def test_soil_confirmation_and_coverage(self):
  h=['序号','面积','TU_RANG_LX','YOU_SHI_SZ','SEN_LIN_LB']
  raw=[['01',.1,'122','706001','011'],['02',.1,'999','706000','11']]
  rows,w=m.prepare(h,raw)
  self.assertEqual(rows[0][27],'122');self.assertEqual(rows[0][16],'706001')
  self.assertTrue(any(x.startswith('Known TU_RANG_LX') for x in w))
  rows,w=m.prepare(h,raw,soil_dictionary='revision-reference')
  self.assertEqual([r[27] for r in rows],['褐土','999'])
  self.assertEqual([r[11] for r in rows],['重点公益林','重点公益林'])
  audit=m.dictionary_coverage(h,raw)
  self.assertEqual(audit['TU_RANG_LX']['requires_source_confirmation'],['122'])
  self.assertEqual(audit['TU_RANG_LX']['unknown_values'],['999'])
  self.assertEqual(audit['YOU_SHI_SZ']['unknown_values'],['706001'])
  self.assertEqual(m.dictionary_coverage(h,raw,'revision-reference')['TU_RANG_LX']['requires_source_confirmation'],[])
  with self.assertRaises(ValueError):m.prepare(h,raw,soil_dictionary='unknown')
 def test_common_code_families(self):
  for field,expected in [('BH_DJ',['Ⅰ','Ⅱ','Ⅲ','Ⅳ']),('LING_ZU',['幼龄林','中龄林','近熟林','成熟林','过熟林'])]:
   idx=next(k for k,v in m.FIELDS.items() if v==field)
   rows,_=m.prepare(['序号','面积',field],[[str(i),.1,str(i)] for i in range(1,len(expected)+1)])
   self.assertEqual([r[idx] for r in rows],expected)
  self.assertEqual([m.MAPS['QI_YUAN'][str(x)] for x in [11,12,13,21,22,23,24]],['天然']*3+['人工']*4)
  self.assertEqual([m.MAPS['LD_QS'][str(x)] for x in [20,21,22,23]],['集体']*4)
 def test_standalone_install_uses_local_dictionary(self):
  with tempfile.TemporaryDirectory() as tmp:
   root=pathlib.Path(tmp)/'portable-skill';shutil.copytree(SKILL,root)
   spec=importlib.util.spec_from_file_location('portable_field',root/'scripts/forestry_field_survey_table.py')
   portable=importlib.util.module_from_spec(spec);spec.loader.exec_module(portable)
   rows,_=portable.prepare(['序号','面积','LIN_ZHONG','YOU_SHI_SZ'],[['001',.1,'251','706000']])
   self.assertEqual(rows[0][6],'001');self.assertEqual(rows[0][16],'杏')
   self.assertEqual(portable.ROOT,root.resolve())
   for p in root.rglob('*'):
    if p.suffix in ('.md','.json','.py','.ps1','.yaml'):
     self.assertNotIn('ink.doro',p.read_text(encoding='utf-8-sig'))
 def test_mapping_sort_and_unknowns(self):
  h=['序号','面积','LIN_BAN','XIAO_BAN','LIN_ZHONG','YOU_SHI_SZ','TU_RANG_LX','PINGJUN_SG','PO_DU','塔号']
  rows,w=m.prepare(h,[['02',.1,'0327','00004','242','410001','122',0,9,'T2'],['01',.2,'0327','00003','111','200000','122',None,4,'T1']])
  self.assertEqual(rows[0][6],'01');self.assertEqual(rows[0][13],'水源林');self.assertEqual(rows[0][16],'油松')
  self.assertEqual(rows[1][4:6],['0327','00004']);self.assertEqual(rows[1][13],'能源林');self.assertEqual(rows[1][16],'410001');self.assertEqual(rows[1][27],'122')
  self.assertEqual(rows[1][18],0);self.assertIsNone(rows[0][18]);self.assertEqual(rows[1][26],9)
  for i in (0,1,2,22,23,29,30,31):self.assertIsNone(rows[1][i])
 def test_unknown_forest_type_stays_code(self):
  rows,w=m.prepare(['序号','面积','LIN_ZHONG'],[['1',.1,'999']])
  self.assertEqual(rows[0][13],'999')
  self.assertIn('Untranslated LIN_ZHONG: 999',w)
 def test_invalid_inputs(self):
  for h,rows in [(['序号','TBMJ'],[['1',2]]),(['序号','面积'],[['1',None]]),(['序号','面积'],[['1',-1]]),(['序号','面积'],[['1','NaN']]),(['序号','面积'],[['1',1],['1',2]]),(['序号','面积','面积'],[['1',1,2]]),(['序号','面积'],[])]:
   with self.subTest(h=h,rows=rows):
    with self.assertRaises(ValueError):m.prepare(h,rows)
 def test_header_normalization_is_whitespace_only(self):
  rows,_=m.prepare([' 序号\n','面 积','DLMC'],[['001',.1,'乔木林地']])
  self.assertEqual(rows[0][6],'001');self.assertEqual(rows[0][7],.1)
  with self.assertRaisesRegex(ValueError,'Required field missing: 序号'):
   m.prepare(['小班编号','面积'],[['001',.1]])
  with self.assertRaisesRegex(ValueError,'Required field missing: 面积'):
   m.prepare(['序号','TBMJ'],[['001',.1]])
 def test_explicit_fields_and_blank_rows(self):
  rows,_=m.prepare(['ID','HA','市'],[[None,None,None],['A',0,'已有市']],id_field='ID',area_field='HA',county='已确认县')
  self.assertEqual(rows[0][0:3],['已有市','已确认县',None]);self.assertEqual(rows[0][7],0)
 def test_default_output(self):
  self.assertEqual(m.default_output(pathlib.Path('/tmp/调查_初表.xls')).name,'调查_外调.xlsx')
 def test_template_classification_and_reference_stock(self):
  h=['序号','面积','DLMC','LIN_ZHONG','SEN_LIN_LB','SHI_QUAN_D','MEI_GQ_XJ']
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
    rows,w=m.prepare(['序号','面积','DLMC','LIN_ZHONG'],[['1',.1,land,forest]])
    self.assertEqual(rows[0][9],expected)
    if '灌木' in land:self.assertEqual(rows[0][8],'一般灌木林地')
    if expected is None:self.assertEqual(rows[0][13],'999');self.assertTrue(any('未匹配林地类型' in x for x in w))
 def test_workflow_observations_are_independent(self):
  headers=['序号','面积','SHI_QUAN_D','LD_QS','LIN_ZHONG','PINGJUN_SG','YU_BI_DU','TU_RANG_LX','PO_XIANG','PO_WEI','PO_DU','TU_CENG_HD','塔号']
  raw=[['02',.1,'20','21','242',5,.15,'122','北','上坡',9,12,'T2'],['01',.2,'10','22','111',8,.7,'123','南','下坡',4,8,'T1']]
  rows,_=m.prepare(headers,raw)
  sheets=m.workflow_rows(['_id']+headers[1:],raw,rows,'临时')
  self.assertEqual([s['name'] for s in sheets],['原始因子表','外业调查表'])
  original,field=[s['rows'] for s in sheets]
  self.assertEqual([len(original[0]),len(field[0])],[35,24])
  self.assertEqual(original[0][12],'国家公益林');self.assertEqual(original[1][12],'地方公益林')
  self.assertEqual(original[1][22],.15);self.assertEqual(original[1][33],'临时')
  for row in field:
   for i in [3,6,7,8,9,10,11,12,14,16,18,20,22]:self.assertIsNone(row[i])
  self.assertEqual([r[19] for r in field],['123','122'])
  self.assertEqual([r[13] for r in field],['南','北'])
  self.assertEqual([r[15] for r in field],['下坡','上坡'])
  self.assertEqual([r[17] for r in field],[4,9])
  self.assertEqual([r[21] for r in field],[8,12])
  self.assertEqual([r[4] for r in field],['集体','集体'])
  self.assertEqual([r[23] for r in field],['T1','T2'])
  field[0][3]='疏林地';self.assertIsNone(original[0][8])
 def test_summary_titles_stay_template_headings(self):
  path=ROOT/'skills/forestry-survey-summary-tables/scripts/summary_tables.py'
  spec=importlib.util.spec_from_file_location('summary_titles',path)
  summary=importlib.util.module_from_spec(spec);spec.loader.exec_module(summary)
  for nature in ['临时','永久']:
   book=summary.read_book(ROOT/'skills/forestry-survey-summary-tables/assets/report-headers.xlsx')
   values={0:'示例市',1:'示例县',2:'示例乡',3:'示例村',6:'001',7:.1,9:'乔木林地',10:'防护林林地',11:'集体',12:'一般商品林',14:'Ⅱ',17:'其他区域',23:.6,24:8.5,25:1.2,26:17,34:nature}
   for index,value in values.items():book['内业处理表']['cells'][f'{summary.col(index+1)}4']={'value':value}
   cfg={'project_name':'跨地区测试工程','use_nature':nature,'ecology':{'mode':'none','confirmed':True},'fee':{'mode':'reference','region':'河北省','urban_planning_area':False,'project_kind':'公共基础设施'}}
   records=summary.prepare(book,cfg);sheets=summary.build(records,cfg)
   self.assertEqual([s['name'] for s in sheets],summary.REPORT_NAMES)
   expected={
    '现状调查表':'建设项目使用林地现状调查表',
    'A.1':'表A.1 使用林地小班因子调查表',
    'B.1':'表B.1 使用林地按林地类型面积与蓄积统计表',
    'B.2':'表B.2 使用林地按现状地类面积与蓄积统计表',
    'B.3':'表B.3 使用林地按森林类别面积与蓄积统计表',
    'B.4':'表B.4 使用林地按林地保护等级面积与蓄积统计表',
    'B.5':'表B.5 使用林地分森林类别按现状地类面积统计表',
    'B.6':'表B.6 使用重点生态区域林地面积统计表',
    'B.7':'表B.7 使用林地森林植被恢复费参考测算表',
   }
   for sheet in sheets:self.assertEqual(sheet['cells']['A1'],expected[sheet['name']])
   self.assertEqual(sheets[0]['cells']['A16'],'建设项目使用林地现状调查表（续）')
 def summary_fixture(self):
  path=ROOT/'skills/forestry-survey-summary-tables/scripts/summary_tables.py'
  spec=importlib.util.spec_from_file_location('summary_review',path)
  summary=importlib.util.module_from_spec(spec);spec.loader.exec_module(summary)
  book=summary.read_book(path.parents[1]/'assets/report-headers.xlsx')
  values={0:'示例市',1:'示例县',2:'示例乡',3:'示例村',4:32,5:167,6:'001',7:.1,9:'乔木林地',10:'防护林林地',11:'集体',12:'一般商品林',14:'Ⅱ',17:'其他区域',23:.6,24:8.5,25:12.15,26:17,27:17,34:'临时'}
  for index,value in values.items():book['内业处理表']['cells'][f'{summary.col(index+1)}4']={'value':value,'formula':None,'style':0}
  cfg={'project_name':'测试工程','use_nature':'临时','ecology':{'mode':'none','confirmed':True},'fee':{'mode':'reference','region':'河北省','urban_planning_area':False,'project_kind':'公共基础设施'}}
  return summary,book,cfg
 def test_summary_rejects_invalid_measurements(self):
  for column,value in [('X',20),('X',1.01),('X',-.1),('X','20%'),('AA',1.5),('AA',-1),('W',float('nan'))]:
   with self.subTest(column=column,value=value):
    summary,book,cfg=self.summary_fixture();book['内业处理表']['cells'][column+'4']['value']=value
    with self.assertRaises(ValueError):summary.prepare(book,cfg)
 def test_summary_ecology_must_agree_with_input(self):
  summary,book,cfg=self.summary_fixture()
  cfg['ecology']={'mode':'records','records':[{'name':'测试保护区','kind':'自然保护区','level':'国家级','parcel_ids':['001']}]}
  with self.assertRaisesRegex(ValueError,'conflicts'):summary.prepare(book,cfg)
  book['内业处理表']['cells']['R4']['value']='自然保护区'
  records=summary.prepare(book,cfg);sheets=summary.build(records,cfg)
  self.assertEqual(next(s for s in sheets if s['name']=='B.6')['checks']['B6'],.1)
  cfg['ecology']={'mode':'none','confirmed':True}
  with self.assertRaises(ValueError):summary.prepare(book,cfg)
 def test_summary_preserves_identifier_padding_and_numeric_roles(self):
  summary,book,cfg=self.summary_fixture();sh=book['内业处理表'];styles=sh['styles']
  formats=styles.find('s:numFmts',summary.NS)
  if formats is None:formats=E.SubElement(styles,'{'+summary.NS['s']+'}numFmts')
  E.SubElement(formats,'{'+summary.NS['s']+'}numFmt',{'numFmtId':'999','formatCode':'00000'})
  xfs=styles.find('s:cellXfs',summary.NS);E.SubElement(xfs,'{'+summary.NS['s']+'}xf',{'numFmtId':'999'})
  sh['cells']['F4']['style']=len(xfs)-1
  records=summary.prepare(book,cfg);sheets=summary.build(records,cfg);a=next(s for s in sheets if s['name']=='A.1')
  effective={addr:fmt for area,fmt in a['formats'] for addr in summary.addresses(area)}
  self.assertEqual(effective['F9'],'00000');self.assertEqual(a['display_checks']['F9'],'00167')
  self.assertEqual(effective['X9'],'0;-0;;@');self.assertEqual(effective['V9'],'0.00;-0.00;;@')
  sh['cells']['AA4']['value']=0
  a=next(s for s in summary.build(summary.prepare(book,cfg),cfg) if s['name']=='A.1')
  self.assertEqual(a['display_checks']['X9'],'')
 def test_summary_permanent_note_excludes_temporary_policy(self):
  summary,book,cfg=self.summary_fixture();cfg['use_nature']='永久';book['内业处理表']['cells']['AI4']['value']='永久'
  sheets=summary.build(summary.prepare(book,cfg),cfg);fee=next(s for s in sheets if s['name']=='B.7')
  self.assertFalse(any('临时使用林地' in str(value) for value in fee['cells'].values()))
  self.assertEqual(fee['cells']['A22'],'')
  self.assertIn('A22:M29',fee['merges'])
  self.assertEqual(fee['style_rows'],{str(row):22+((row-22)%4) for row in range(22,30)})
  self.assertEqual(fee['last'],29)
  self.assertIn(str(fee['last']),fee['style_rows'])

 def test_summary_manual_continuation_bodies_are_blank(self):
  summary,book,cfg=self.summary_fixture()
  sheets=summary.build(summary.prepare(book,cfg),cfg)
  status=next(s for s in sheets if s['name']=='现状调查表')
  for address in ['A25','A27','A29','A31','A33','A35']:
   self.assertEqual(status['cells'][address],'')
  self.assertEqual(status['cells']['A37'],'编制单位意见（单位印章）：')
  self.assertEqual(status['alignments'],[
   ['A25:S25','left','center'],['A27:S27','left','center'],
   ['A29:S29','left','center'],['A31:S31','left','center'],
   ['A33:S33','left','center'],['A35:S35','left','center'],
   ['A37:S37','left','center'],
  ])

 def test_summary_ecology_none_is_blank_in_a1(self):
  summary,book,cfg=self.summary_fixture()
  sheets=summary.build(summary.prepare(book,cfg),cfg)
  detail=next(s for s in sheets if s['name']=='A.1')
  ecology=detail['cells']['P9']
  self.assertIn('""',ecology)
  self.assertNotIn('),"不涉及",',ecology)
  b6=next(s for s in sheets if s['name']=='B.6')
  self.assertIn('COUNTIFS',b6['cells']['A6'])

 def test_summary_b5_matches_current_template_rows(self):
  summary,book,cfg=self.summary_fixture()
  sheets=summary.build(summary.prepare(book,cfg),cfg)
  b5=next(s for s in sheets if s['name']=='B.5')
  self.assertEqual((b5['head'],b5['cols'],b5['last']),(5,15,19))
  self.assertIn('A1:O1',b5['merges'])
  self.assertIn('A2:O2',b5['merges'])
  self.assertIn('A3:A4',b5['merges'])
  self.assertIn('A6:A12',b5['merges'])
  self.assertIn('A13:A19',b5['merges'])
  self.assertNotIn('A1:O2',b5['merges'])
 def test_summary_layout_uses_template_without_print_settings(self):
  summary,book,cfg=self.summary_fixture();sheets=summary.build(summary.prepare(book,cfg),cfg)
  book['现状调查表']['cells']['A1']['value']='已核定模板标题'
  summary.report_layout(sheets,book)
  self.assertEqual(sheets[0]['cells']['A1'],'已核定模板标题')
  for sheet in sheets:
   for key in ('manual_breaks','print_titles','print_area'):self.assertNotIn(key,sheet)
 def test_summary_restores_column_defaults_with_new_style_indexes(self):
  summary,book,cfg=self.summary_fixture();template=summary.ROOT/'assets/report-headers.xlsx'
  with tempfile.TemporaryDirectory() as directory:
   source=pathlib.Path(directory)/'source.xlsx';candidate=pathlib.Path(directory)/'candidate.xlsx'
   candidate.write_bytes(template.read_bytes())
   with zipfile.ZipFile(template) as z:parts={n:z.read(n) for n in z.namelist()}
   from xml.dom import minidom
   styles=minidom.parseString(parts['xl/styles.xml']);uri=summary.NS['s']
   formats=styles.getElementsByTagNameNS(uri,'numFmts')[0]
   fmt=styles.createElementNS(uri,'numFmt');fmt.setAttribute('numFmtId','999');fmt.setAttribute('formatCode','0.00000')
   formats.appendChild(fmt);formats.setAttribute('count',str(len(formats.getElementsByTagNameNS(uri,'numFmt'))))
   xfs=styles.getElementsByTagNameNS(uri,'cellXfs')[0];nodes=xfs.getElementsByTagNameNS(uri,'xf');index=len(nodes)
   xf=nodes[0].cloneNode(True);xf.setAttribute('numFmtId','999');xfs.appendChild(xf);xfs.setAttribute('count',str(index+1))
   part=book['内业处理表']['part'];sheet=minidom.parseString(parts[part])
   sheet.getElementsByTagNameNS(uri,'col')[0].setAttribute('style',str(index))
   parts[part]=sheet.toxml(encoding='utf-8');parts['xl/styles.xml']=styles.toxml(encoding='utf-8')
   with zipfile.ZipFile(source,'w') as z:
    for name,data in parts.items():z.writestr(name,data)
   with zipfile.ZipFile(candidate) as z:damaged={n:z.read(n) for n in z.namelist()}
   cell_sheet=minidom.parseString(damaged[part])
   for cell in cell_sheet.getElementsByTagNameNS(uri,'c'):
    if cell.getAttribute('r')=='A1':cell.setAttribute('s','0')
   damaged[part]=cell_sheet.toxml(encoding='utf-8')
   with zipfile.ZipFile(candidate,'w') as z:
    for name,data in damaged.items():z.writestr(name,data)
   summary.restore_retained_styles(source,candidate)
   summary.verify_preserved(summary.read_book(source),summary.read_book(candidate),summary.INPUT_NAMES)
  def test_summary_preserves_retained_input_without_print_metadata_checks(self):
   summary,book,cfg=self.summary_fixture()
   template=summary.ROOT/'assets/report-headers.xlsx'
   with tempfile.TemporaryDirectory() as directory:
    candidate=pathlib.Path(directory)/'candidate.xlsx';candidate.write_bytes(template.read_bytes())
    before=summary.read_book(candidate)
    after=summary.read_book(candidate)
    summary.verify_preserved(before,after,summary.INPUT_NAMES)
  def test_summary_report_fonts_drop_color(self):
   summary,book,cfg=self.summary_fixture()
   template=summary.ROOT/'assets/report-headers.xlsx'
   payload=summary.build(summary.prepare(book,cfg),cfg)
   with tempfile.TemporaryDirectory() as directory:
    candidate=pathlib.Path(directory)/'candidate.xlsx'
    with zipfile.ZipFile(template) as z:parts={n:z.read(n) for n in z.namelist()}
    from xml.dom import minidom
    styles=minidom.parseString(parts['xl/styles.xml'])
    for color in styles.getElementsByTagNameNS(summary.NS['s'],'color'):
     if color.parentNode.localName=='font':
      for key in list(color.attributes.keys()):color.removeAttribute(key)
      color.setAttribute('rgb','FFFFFFFF')
    parts['xl/styles.xml']=styles.toxml(encoding='utf-8')
    with zipfile.ZipFile(candidate,'w') as z:
     for name,data in parts.items():z.writestr(name,data)
    before=summary.read_book(candidate)
    # Template-sized fixture; all roles, including B.5, must retain input styles.
    payload=[{'name':n,'cols':15 if n=='B.5' else 1,'head':2,'last':2} for n in summary.REPORT_NAMES]
    self.assertGreater(summary.restore_style_components(template,candidate,'font',payload),0)
    after=summary.read_book(candidate);summary.verify_preserved(before,after,summary.INPUT_NAMES)
    for name in summary.REPORT_NAMES:
     a=after[name];b=book[name]
     def font(sh):
      xf=sh['styles'].find('s:cellXfs',summary.NS)[sh['cells']['A1']['style']]
      return E.tostring(sh['styles'].find('s:fonts',summary.NS)[int(xf.attrib.get('fontId',0))])
     def colorless(value):
      node=E.fromstring(value)
      for child in list(node):
       if child.tag=='{'+summary.NS['s']+'}color':node.remove(child)
      return E.tostring(node)
     self.assertEqual(colorless(font(a)),colorless(font(b)))
     self.assertIsNone(E.fromstring(font(a)).find('s:color',summary.NS))
   self.assertEqual(summary.restore_style_components(template,candidate,'font',payload),0)
  def test_summary_preservation_rejects_retained_sheet_metadata_drift(self):
   from copy import deepcopy
   summary,book,cfg=self.summary_fixture()
   for field in ['names','state']:
    with self.subTest(field=field):
     other=deepcopy(book);sh=other['内业处理表']
     if field=='names':sh['names']['_xlnm.Print_Area']='changed'
     elif field=='state':sh['state']='hidden'
     with self.assertRaisesRegex(ValueError,'setting changed'):summary.verify_preserved(book,other,summary.INPUT_NAMES)
 def test_summary_timeout_still_runs_owned_session_cleanup(self):
  from unittest.mock import patch
  import subprocess
  summary,book,cfg=self.summary_fixture()
  with patch.object(summary.subprocess,'run',side_effect=[subprocess.TimeoutExpired('Excel',1),None]) as run:
   with self.assertRaises(subprocess.TimeoutExpired):summary.run_excel(pathlib.Path('job.json'),timeout=1)
   self.assertEqual(run.call_count,2)
   self.assertEqual(run.call_args_list[1].args[0][-1],'-CleanupOnly')
 def test_summary_publication_refuses_changed_source_and_existing_output(self):
  import json
  summary,book,cfg=self.summary_fixture()
  with tempfile.TemporaryDirectory() as directory:
   base=pathlib.Path(directory);source=base/'source.xlsx';candidate=base/'candidate.xlsx';manifest=base/'stage.json'
   source.write_bytes(b'original');candidate.write_bytes(b'validated')
   metadata={'source':str(source),'output':str(source),'source_hash':summary.digest(source),'candidate_hash':summary.digest(candidate),'records':1,'area_ha':.1,'volume_m3':1.2}
   manifest.write_text(json.dumps(metadata),encoding='utf-8');source.write_bytes(b'user edit')
   with self.assertRaisesRegex(ValueError,'changed since validation'):summary.publish(manifest)
   self.assertEqual(source.read_bytes(),b'user edit')
   source.write_bytes(b'original');metadata['output']=str(candidate);manifest.write_text(json.dumps(metadata),encoding='utf-8')
   with self.assertRaises(FileExistsError):summary.publish(manifest)
   self.assertEqual(candidate.read_bytes(),b'validated')
 def test_summary_restores_fixed_template_borders_only(self):
  summary,book,cfg=self.summary_fixture()
  template=summary.ROOT/'assets/report-headers.xlsx'
  with tempfile.TemporaryDirectory() as directory:
   candidate=pathlib.Path(directory)/'candidate.xlsx'
   with zipfile.ZipFile(template) as z:parts={n:z.read(n) for n in z.namelist()}
   from xml.dom import minidom
   styles=minidom.parseString(parts['xl/styles.xml'])
   for border in styles.getElementsByTagNameNS(summary.NS['s'],'border'):
    for child in list(border.childNodes):border.removeChild(child)
   for fill in styles.getElementsByTagNameNS(summary.NS['s'],'patternFill'):
    fill.setAttribute('patternType','solid')
   parts['xl/styles.xml']=styles.toxml(encoding='utf-8')
   with zipfile.ZipFile(candidate,'w') as z:
    for name,data in parts.items():z.writestr(name,data)
   before=summary.read_book(candidate);reference=summary.read_book(template)
   with self.assertRaisesRegex(ValueError,'border changed'):
    summary.verify_report_borders(before['现状调查表'],reference['现状调查表'],19,35)
   payload=[{'name':'现状调查表','cols':19,'head':35,'last':35}]
   self.assertGreater(summary.restore_style_components(template,candidate,'border',payload),0)
   self.assertGreater(summary.restore_style_components(template,candidate,'fill',payload),0)
   after=summary.read_book(candidate)
   summary.verify_report_borders(after['现状调查表'],reference['现状调查表'],19,35)
   summary.verify_preserved(before,after,summary.INPUT_NAMES)
   self.assertEqual(summary.restore_style_components(template,candidate,'border',payload),0)
   self.assertEqual(summary.restore_style_components(template,candidate,'fill',payload),0)
 def test_summary_restores_only_input_title_fonts(self):
  from xml.dom import minidom
  path=ROOT/'skills/forestry-survey-summary-tables/scripts/summary_tables.py'
  spec=importlib.util.spec_from_file_location('summary_fonts',path)
  summary=importlib.util.module_from_spec(spec);spec.loader.exec_module(summary)
  uri=summary.NS['s']
  with tempfile.TemporaryDirectory() as directory:
   source=pathlib.Path(directory)/'source.xlsx';candidate=pathlib.Path(directory)/'candidate.xlsx'
   with zipfile.ZipFile(path.parents[1]/'assets/report-headers.xlsx') as archive:
    parts={n:archive.read(n) for n in archive.namelist()}
   styles=minidom.parseString(parts['xl/styles.xml'])
   fonts=styles.getElementsByTagNameNS(uri,'fonts')[0];xfs=styles.getElementsByTagNameNS(uri,'cellXfs')[0]
   # Different source fonts prove restoration is source-driven, not hardcoded 11pt.
   for index,size in enumerate([11,13,15],1):
    sheet=minidom.parseString(parts[f'xl/worksheets/sheet{index}.xml'])
    font=styles.createElementNS(uri,'font')
    for tag,value in [('name','宋体'),('sz',str(size)),('b','1'),('i','1' if index==2 else '0')]:
     node=styles.createElementNS(uri,tag);node.setAttribute('val',value);font.appendChild(node)
    fonts.appendChild(font);font_id=len(fonts.getElementsByTagNameNS(uri,'font'))-1
    for cell in sheet.getElementsByTagNameNS(uri,'c'):
     if not cell.getAttribute('r').rstrip('0123456789')+'1'==cell.getAttribute('r'):continue
     xf=xfs.getElementsByTagNameNS(uri,'xf')[int(cell.getAttribute('s') or '0')].cloneNode(True)
     xf.setAttribute('fontId',str(font_id));xfs.appendChild(xf)
     cell.setAttribute('s',str(len(xfs.getElementsByTagNameNS(uri,'xf'))-1))
    # Keep namespace declarations that a naive ET rewrite could discard.
    sheet.documentElement.setAttribute('xmlns:mc','http://schemas.openxmlformats.org/markup-compatibility/2006')
    sheet.documentElement.setAttribute('xmlns:testext','urn:title-font-test')
    sheet.documentElement.setAttribute('mc:Ignorable','testext')
    parts[f'xl/worksheets/sheet{index}.xml']=sheet.toxml(encoding='utf-8')
   fonts.setAttribute('count',str(len(fonts.getElementsByTagNameNS(uri,'font'))))
   xfs.setAttribute('count',str(len(xfs.getElementsByTagNameNS(uri,'xf'))))
   parts['xl/styles.xml']=styles.toxml(encoding='utf-8')
   def write(destination,items):
    with zipfile.ZipFile(destination,'w',zipfile.ZIP_DEFLATED) as archive:
     for name,value in items.items():archive.writestr(name,value)
   write(source,parts)
   # Reproduce font drift throughout the merged title, not merely A1.
   changed=dict(parts)
   styles=minidom.parseString(parts['xl/styles.xml'])
   fonts=styles.getElementsByTagNameNS(uri,'fonts')[0]
   for font in list(fonts.getElementsByTagNameNS(uri,'font'))[-3:]:
    font.getElementsByTagNameNS(uri,'sz')[0].setAttribute('val','8')
    font.getElementsByTagNameNS(uri,'b')[0].setAttribute('val','0')
   changed['xl/styles.xml']=styles.toxml(encoding='utf-8');write(candidate,changed)
   before=summary.read_book(source)
   with self.assertRaisesRegex(ValueError,'Preserved style changed'):
    summary.verify_preserved(before,summary.read_book(candidate),summary.INPUT_NAMES)
   self.assertGreater(summary.restore_input_title_fonts(source,candidate),3)
   after=summary.read_book(candidate)
   summary.verify_preserved(before,after,list(before))
   with zipfile.ZipFile(candidate) as archive:
    for name,data in parts.items():
     if name not in ['xl/styles.xml']+[f'xl/worksheets/sheet{i}.xml' for i in range(1,4)]:
      self.assertEqual(archive.read(name),data)
    self.assertIn(b'xmlns:testext="urn:title-font-test"',archive.read('xl/worksheets/sheet1.xml'))
   digest=summary.digest(candidate)
   self.assertEqual(summary.restore_input_title_fonts(source,candidate),0)
   self.assertEqual(summary.digest(candidate),digest)
   with self.assertRaisesRegex(ValueError,'separate candidate'):
    summary.restore_input_title_fonts(source,source)
   # A body change must still block delivery; title restoration is not a bypass.
   after['内业处理表']['cells']['A4']['value']='unexpected edit'
   with self.assertRaisesRegex(ValueError,'Preserved cell changed'):
    summary.verify_preserved(before,after,summary.INPUT_NAMES)
 def test_template_clean_and_styled(self):
  ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
  with zipfile.ZipFile(SKILL/'assets/A.1.xlsx') as z:
   self.assertNotIn(b'absPath',z.read('xl/workbook.xml'))
   wb=E.fromstring(z.read('xl/workbook.xml'));self.assertEqual([x.get('name') for x in wb.find('m:sheets',ns)],['原始因子表','外业调查表'])
   self.assertFalse(any('externalLink' in n or 'vbaProject' in n for n in z.namelist()))
   for i in (1,2):
    sh=E.fromstring(z.read(f'xl/worksheets/sheet{i}.xml'))
    for cell in sh.findall('.//m:sheetData/m:row/m:c',ns):
     row=int(''.join(c for c in cell.get('r') if c.isdigit()))
     if row>=4:self.assertIsNone(cell.find('m:v',ns));self.assertIsNone(cell.find('m:f',ns))
    self.assertNotIn(b'absPath',z.read(f'xl/worksheets/sheet{i}.xml'))
   source_sheet=z.read('xl/worksheets/sheet1.xml')
   self.assertIn('临时,永久'.encode('utf-8'),source_sheet)
   self.assertIn(b'G4:G1048576',source_sheet)
   self.assertIn(b'J4:J1048576',source_sheet)
   strings=E.fromstring(z.read('xl/sharedStrings.xml'));texts=[''.join(x.itertext()) for x in strings]
   normalize=lambda v:''.join(v.split()).replace('²','2').replace('（','(').replace('）',')')
   expected_headers={
    1:['市','县','乡','村','林班号(村小组)','小班号林地落界','使用林地序号','面积hm2','原始地类','林地类型','林地权属','森林类别','公益林事权','林地保护等级','林种','起源','重点生态区域','原优势树种','龄组(产期)','林龄','平均树高m','平均胸径cm','郁闭度(覆盖度)','每公顷蓄积m3','小班蓄积m3','株数','坡向','坡位','坡度','土壤类型','土层厚度cm','建设内容','工程类别','使用林地性质','备注'],
    2:['使用林地序号','面积hm2','原始地类','现状地类','林地权属','原优势树种','优势树种','龄组(产期)','林龄','平均树高m','平均胸径cm','郁闭度(覆盖度)','株数','坡向','坡向','坡位','坡位','坡度','坡度','土壤类型','土壤类型','土层厚度','土层厚度','备注']}
   for sheet_index,expected in expected_headers.items():
    sheet=E.fromstring(z.read(f'xl/worksheets/sheet{sheet_index}.xml'));headers=[]
    for c in sheet.findall('.//m:row[@r="3"]/m:c',ns):
     v=c.find('m:v',ns)
     if v is not None:headers.append(texts[int(v.text)])
    self.assertEqual(list(map(normalize,headers)),expected)
   self.assertNotIn('旧编号',texts)
   self.assertFalse(any('三家' in x or '承德' in x for x in texts))
if __name__=='__main__':unittest.main()
