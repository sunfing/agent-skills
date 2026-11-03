"""Business invariants for ownership certificates; Office integration is checked separately."""
import importlib.util
from copy import deepcopy
import json
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / 'skills/forestry-tree-ownership-certificate/scripts/ownership_certificate.py'
spec = importlib.util.spec_from_file_location('ownership', SCRIPT)
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)


def config():
    return {
        'project_name': '示例110千伏工程', 'project_name_basis': '示例批复第1页',
        'applicant': '示例电力公司', 'applicant_basis': '示例申请材料',
        'use_nature': '永久', 'project_dir': 'D:/example',
        'source': {'path': 'source.xlsx', 'sheet': 'A.1', 'data_rows': [4, 5],
                   'columns': {'id': 'A', 'area': 'B', 'stock': 'C', 'trees': 'D'},
                   'area_basis': '占用面积公顷', 'stock_basis': '地块总蓄积立方米',
                   'trees_basis': '地块总株数', 'confirmed': True},
        'groups': [{'city': '某市', 'county': '某县', 'township': '某镇', 'village': '某村',
                    'owner': '某村集体', 'kind': '集体', 'rows': [4, 5],
                    'ownership_basis': '已核实登记材料', 'ownership_confirmed': True,
                    'undisputed_confirmed': True, 'trees_mode': 'include'}],
    }


def records():
    return [{'row': 4, 'id': '1', 'area': '.01234', 'stock': '.25', 'trees': 2},
            {'row': 5, 'id': '2', 'area': '.00001', 'stock': '.25', 'trees': 0}]


class OwnershipTests(unittest.TestCase):
    def test_default_sheet_and_explicit_override(self):
        c = config(); del c['source']['sheet']
        m.validate_config(c)
        self.assertEqual(c['source']['sheet'], 'A.1')
        c['source']['sheet'] = '用户指定工作表'
        m.validate_config(c)
        self.assertEqual(c['source']['sheet'], '用户指定工作表')
        c['source']['sheet'] = ''
        with self.assertRaises(ValueError): m.validate_config(c)

    def test_project_root_and_output_for_both_use_natures(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / '示例项目'
            for nature in ('永久', '临时'):
                c = config(); del c['project_dir']; c['use_nature'] = nature
                c['source']['path'] = str(root / '05_报告表格与图件' / '表格' / nature / '调查表.xlsx')
                m.validate_config(c)
                self.assertEqual(Path(c['project_dir']), root.resolve())
                self.assertEqual(m.default_target(c), root.resolve() / '06_公示查验与报批' / nature / '林木权属证明' / f'林木权属证明（{nature}）.docx')
            for path in (root / '调查表.xlsx', root / '05_报告表格与图件' / '05_报告表格与图件' / '调查表.xlsx'):
                with self.assertRaisesRegex(ValueError, 'project_dir'):
                    m.infer_project_dir(path)
            c = config(); c['project_dir'] = ''
            with self.assertRaises(ValueError): m.validate_config(c)

    def test_missing_required_fields_are_reported_for_default_sheet(self):
        for key in ('id', 'area', 'stock'):
            c = config(); del c['source']['sheet']; del c['source']['columns'][key]
            with self.subTest(key=key), self.assertRaisesRegex(ValueError, f'A.1: {key}'):
                m.validate_config(c)

    def test_sum_before_half_up_and_integer_tree_count(self):
        g = m.prepare(config(), records())[0]
        self.assertEqual((g['area'], g['stock'], g['trees']), ('0.0124', '1', '2'))
        self.assertEqual(g['stock_raw'], '0.50')
        self.assertIn('某村集体所有', g['body'])
        self.assertNotIn('我单位', g['body'])

    def test_same_number_in_other_use_nature_is_independent(self):
        c = config()
        first = m.prepare(c, records())
        c['use_nature'] = '临时'
        second = m.prepare(c, records())
        self.assertEqual(first[0]['ids'], second[0]['ids'])
        self.assertIn('临时使用林地', second[0]['body'])

    def test_no_silent_missing_or_fractional_counts(self):
        for field, value in [('stock', None), ('area', None), ('area', 0),
                             ('trees', None), ('trees', 1.2), ('stock', 'NaN'), ('trees', True)]:
            with self.subTest(field=field, value=value):
                rows = records(); rows[0][field] = value
                with self.assertRaises(ValueError):
                    m.prepare(config(), rows)

    def test_omit_counts_requires_explicit_basis(self):
        c = config(); c['groups'][0]['trees_mode'] = 'omit'
        with self.assertRaises(ValueError):
            m.prepare(c, records())
        c['groups'][0]['trees_omit_basis'] = '已确认缺少株数且同意省略'
        rows = records(); rows[0]['trees'] = None
        g = m.prepare(c, rows)[0]
        self.assertIsNone(g['trees'])
        self.assertNotIn('株数', g['body'])

    def test_group_coverage_and_ownership_confirmation(self):
        for change in ('missing', 'overlap', 'dispute', 'ownership', 'source'):
            c = config()
            if change == 'missing': c['groups'][0]['rows'] = [4]
            if change == 'overlap': c['groups'].append(deepcopy(c['groups'][0]))
            if change == 'dispute': c['groups'][0]['undisputed_confirmed'] = False
            if change == 'ownership': c['groups'][0]['ownership_confirmed'] = False
            if change == 'source': c['source']['confirmed'] = False
            with self.subTest(change=change), self.assertRaises(ValueError):
                m.prepare(c, records())

    def test_different_owners_in_same_village_remain_separate(self):
        c = config(); c['groups'][0]['rows'] = [4]
        g = deepcopy(c['groups'][0]); g['owner'] = '另一个已确认权利人'; g['rows'] = [5]
        c['groups'].append(g)
        self.assertEqual(len(m.prepare(c, records())), 2)

    def test_duplicate_ids_and_extracted_scope_stop(self):
        rows = records(); rows[1]['id'] = '1'
        with self.assertRaises(ValueError): m.prepare(config(), rows)
        with self.assertRaises(ValueError): m.prepare(config(), records()[:1])
        rows = records(); rows[0]['id'] = None
        with self.assertRaises(ValueError): m.prepare(config(), rows)

    def test_invalid_excel_column_and_numeric_names(self):
        c = config(); c['source']['columns']['stock'] = 'XFE'
        with self.assertRaises(ValueError): m.prepare(c, records())
        c = config(); c['groups'][0].update(township='第1镇', village='第2村', owner='第2村集体')
        body = m.prepare(c, records())[0]['body']
        self.assertIn('第 1 镇第 2 村', body)
        self.assertIn('第 2 村集体所有', body)

    def test_compress_only_true_numeric_ranges(self):
        self.assertEqual(m.compress_ids(['5', '1', '2', '4']), '1-2、4-5')
        self.assertEqual(m.compress_ids(['01', '02', '03']), '01、02、03')
        self.assertEqual(m.compress_ids(['A1', 'A2']), 'A1、A2')

    def test_docx_signing_roles_and_identical_dates(self):
        c = config(); c['groups'][0]['rows'] = [4]
        g = deepcopy(c['groups'][0]); g.update(kind='国有', owner='某林场', village='某林场', township='', rows=[5])
        c['groups'].append(g)
        certs = m.prepare(c, records())
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / 'candidate.docx'; m.write_docx(path, c, certs)
            with zipfile.ZipFile(path) as z:
                doc = ET.fromstring(z.read('word/document.xml'))
            ns = {'w': m.W}
            tables = doc.findall('.//w:tbl', ns)
            self.assertEqual(len(tables), 2)
            for table in tables:
                cells = table.findall('.//w:tc', ns)
                dates = [''.join(cell.findall('w:p', ns)[-1].itertext()) for cell in cells]
                self.assertEqual(dates, ['　　　　年　　月　　日'] * 2)
                self.assertTrue(all(x.get(f'{{{m.W}}}val') == 'nil'
                                    for x in table.findall('w:tblPr/w:tblBorders/*', ns)))
            text = ''.join(doc.itertext())
            for expected in ('村委会盖章：', '乡镇政府盖章：', '林场盖章：', '主管单位盖章：', '第一阶段', '第二阶段', '正式报送时不装入'):
                self.assertIn(expected, text)
            self.assertEqual(len(doc.findall('.//w:pageBreakBefore', ns)), 2)

    def test_publish_detects_source_candidate_and_target_changes(self):
        for changed in ('source', 'candidate', 'target'):
            with self.subTest(changed=changed), tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp); source = root/'source'; source.write_bytes(b'source')
                candidate = root/'candidate.docx'; candidate.write_bytes(b'candidate')
                target = root/'target.docx'; target.write_bytes(b'original')
                stage = {'inputs': [{'path': str(source), 'sha256': m.digest(source)}],
                         'candidate': str(candidate), 'candidate_sha256': m.digest(candidate),
                         'target': str(target), 'target_sha256': m.digest(target), 'replace': True}
                stage_path = root/'stage.json'; m.write_json(stage_path, stage)
                {'source': source, 'candidate': candidate, 'target': target}[changed].write_bytes(b'changed')
                before = target.read_bytes()
                with self.assertRaises(ValueError): m.publish(stage_path)
                self.assertEqual(target.read_bytes(), before)

    def test_publish_new_and_replace(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp); candidate = root/'candidate.docx'; candidate.write_bytes(b'one')
            target = root/'out'/'target.docx'
            stage = {'inputs': [], 'candidate': str(candidate), 'candidate_sha256': m.digest(candidate),
                     'target': str(target), 'target_sha256': None, 'replace': False}
            stage_path = root/'stage.json'; m.write_json(stage_path, stage)
            m.publish(stage_path); self.assertEqual(target.read_bytes(), b'one')
            stage.update(target_sha256=m.digest(target), replace=True)
            candidate.write_bytes(b'two'); stage['candidate_sha256'] = m.digest(candidate)
            m.write_json(stage_path, stage); m.publish(stage_path)
            self.assertEqual(target.read_bytes(), b'two')


if __name__ == '__main__':
    unittest.main()
