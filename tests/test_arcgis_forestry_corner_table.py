import importlib.util
import io
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "skills" / "arcgis-forestry-corner-table" / "scripts" / "arcgis_corner_table.py"
spec = importlib.util.spec_from_file_location("corner", SCRIPT)
corner = importlib.util.module_from_spec(spec)
sys.modules["corner"] = corner
assert spec.loader is not None
spec.loader.exec_module(corner)

class CoreTests(unittest.TestCase):
    def p(self,row,parcel,x,y,area=0.1):
        return corner.PointRow(row,parcel,area,"乔木林地","一般商品林","Ⅳ",x,y)

    def test_zone39(self):
        self.assertEqual(corner.detect_zone([self.p(2,1,39603020.0,4573000.0)]),39)

    def test_unbanded_rejected(self):
        with self.assertRaises(corner.ConversionError):
            corner.detect_zone([self.p(2,1,603020.0,4573000.0)])

    def test_two_rings_restart_j1(self):
        pts=[
            self.p(2,1,39603000,4573000),self.p(3,1,39603010,4573000),
            self.p(4,1,39603010,4573010),self.p(5,1,39603000,4573000),
            self.p(6,1,39603100,4573100),self.p(7,1,39603110,4573100),
            self.p(8,1,39603110,4573110),self.p(9,1,39603100,4573100),
        ]
        out,summary=corner.build_output_rows(pts,0.001)
        self.assertEqual(summary[1]["rings"],2)
        self.assertEqual([r.corner for r in out[:4]],["J1","J2","J3","J1"])
        self.assertEqual([r.corner for r in out[4:]],["J1","J2","J3","J1"])

    def test_unclosed_rejected(self):
        with self.assertRaises(corner.ConversionError):
            corner.build_output_rows([
                self.p(2,1,39603000,4573000),self.p(3,1,39603010,4573000),
                self.p(4,1,39603010,4573010)],0.001)

    def test_standardization(self):
        self.assertEqual(corner.std_forest("012"),"一般公益林")
        self.assertEqual(corner.std_forest("022"),"一般商品林")
        self.assertEqual(corner.std_grade(3),"Ⅲ")
        self.assertEqual(corner.std_grade("IV"),"Ⅳ")

    def test_check_only_csv(self):
        data="序号,面积,地类,森林类,保护等,POINT_X,POINT_Y\n" \
             "1,0.1,乔木林地,022,4,39603000,4573000\n" \
             "1,0.1,乔木林地,022,4,39603010,4573000\n" \
             "1,0.1,乔木林地,022,4,39603010,4573010\n" \
             "1,0.1,乔木林地,022,4,39603000,4573000\n"
        with tempfile.TemporaryDirectory() as td:
            p=Path(td)/"zone39.csv"; p.write_text(data,encoding="utf-8-sig")
            buf=io.StringIO()
            with redirect_stdout(buf):
                rc=corner.main([str(p),"--check-only","--expected-zone","39"])
            self.assertEqual(rc,0)
            self.assertIn("Zone 39",buf.getvalue())
            self.assertIn("独立闭合面环：1",buf.getvalue())

if __name__=="__main__":
    unittest.main()
