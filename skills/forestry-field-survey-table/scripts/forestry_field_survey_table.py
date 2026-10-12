"""Prepare the fixed A.1 field survey workbook using desktop Excel (stdlib only)."""
import argparse
from decimal import Decimal, InvalidOperation
import hashlib
import json
import os
from pathlib import Path
import subprocess
import tempfile

ROOT = Path(__file__).resolve().parents[1]
MAPS = {
    "LD_QS": {"10": "国有", "20": "集体"},
    "SEN_LIN_LB": {"012": "一般公益林", "022": "一般商品林"},
    "BH_DJ": {"3": "Ⅲ", "4": "Ⅳ"},
    "LIN_ZHONG": {"111": "水源林", "112": "水保林", "233": "用材林"},
    "QI_YUAN": {"11": "天然", "21": "人工"},
    "YOU_SHI_SZ": {"200000": "油松", "465000": "刺槐", "490000": "其他硬阔类", "590000": "其他软阔类", "999000": "其他灌木"},
    "LING_ZU": {"1": "幼龄林", "4": "成熟林"},
    "PO_XIANG": dict(zip(map(str, range(1, 10)), ["北", "东北", "东", "东南", "南", "西南", "西", "西北", "无坡向"])),
    "PO_WEI": dict(zip(map(str, range(1, 7)), ["脊", "上", "中", "下", "谷", "平地"])),
}
FIELDS = {0: "市", 1: "县", 2: "乡", 3: "ZLDWMC", 4: "LIN_BAN", 5: "XIAO_BAN", 6: "小班编", 7: "面积", 8: "DLMC", 10: "LD_QS", 11: "SEN_LIN_LB", 12: "BH_DJ", 13: "LIN_ZHONG", 14: "QI_YUAN", 15: "STQWMC", 16: "YOU_SHI_SZ", 17: "LING_ZU", 18: "PINGJUN_SG", 19: "PINGJUN_XJ", 20: "YU_BI_DU", 21: "MEI_GQ_XJ", 24: "PO_XIANG", 25: "PO_WEI", 26: "PO_DU", 27: "TU_RANG_LX", 28: "TU_CENG_HD"}
NUMERIC = {18, 19, 20, 21, 26, 28}

def clean(v):
    return v.strip() if isinstance(v, str) else v

def code(v):
    v = clean(v)
    if v is None or v == "":
        return None
    return str(int(v)) if isinstance(v, (float, int)) and float(v).is_integer() else str(v)

def number(v, field):
    if v is None or v == "":
        return None
    try:
        n = Decimal(str(v))
    except InvalidOperation as exc:
        raise ValueError(f"Non-numeric {field}: {v}") from exc
    if not n.is_finite() or n < 0:
        raise ValueError(f"Invalid {field}: {v}")
    return float(n)

def prepare(headers, source_rows, *, area_field="面积", id_field="小班编", city=None, county=None, township=None):
    headers = [clean(h) for h in headers]
    present = [h for h in headers if h]
    if len(set(present)) != len(present):
        raise ValueError("Duplicate source headers")
    for field in (area_field, id_field):
        if field not in headers:
            raise ValueError(f"Required field missing: {field}")
    result, seen, warnings = [], set(), set()
    for values in source_rows:
        if all(v is None or clean(v) == "" for v in values):
            continue
        r = dict(zip(headers, map(clean, values)))
        ident = code(r.get(id_field))
        if not ident or ident in seen:
            raise ValueError(f"Missing or duplicate small-parcel identifier: {ident}")
        seen.add(ident)
        area = number(r.get(area_field), area_field)
        if area is None:
            raise ValueError(f"Missing area for {ident}")
        row = [None] * 33
        for col, field in FIELDS.items():
            v = r.get(field)
            if v is None or v == "":
                if not (col in (0, 1, 2) and (city, county, township)[col]):
                    warnings.add(f"Missing input: {field}")
                continue
            if col in NUMERIC:
                row[col] = number(v, field)
            elif field in MAPS:
                raw = code(v)
                row[col] = MAPS[field].get(raw, raw)
                if raw not in MAPS[field]:
                    warnings.add(f"Untranslated {field}: {raw}")
            else:
                row[col] = code(v)
        row[6], row[7] = ident, area
        for col, value in enumerate((city, county, township)):
            if value:
                row[col] = value
        tower = code(r.get("塔号"))
        row[32] = tower  # compact enough for the fixed Remarks column
        if code(r.get("DI_LEI")) == "030701":
            warnings.add("DI_LEI 030701: verify classification against the source dictionary; DLMC retained")
        result.append(row)
    if not result:
        raise ValueError("No source records")
    # Numeric order when all identifiers are digits; retain their original text.
    if all(row[6].isdigit() for row in result):
        result.sort(key=lambda row: (int(row[6]), row[6]))
    return result, sorted(warnings)

def default_output(source):
    stem = source.stem
    return source.with_name((stem[:-3] if stem.endswith("_初表") else stem) + "_外调.xlsx")

def excel(job, tmp):
    request = Path(tmp) / "job.json"
    request.write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8-sig")
    subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(ROOT / "scripts/excel_io.ps1"), "-Job", str(request)], check=True, timeout=180)

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("input", type=Path)
    ap.add_argument("--sheet", help="Required when the input has multiple worksheets")
    ap.add_argument("--output", type=Path)
    ap.add_argument("--id-field", default="小班编")
    ap.add_argument("--area-field", default="面积", help="Confirmed occupied area in hectares, never total original polygon area")
    ap.add_argument("--city")
    ap.add_argument("--county")
    ap.add_argument("--township")
    ap.add_argument("--check-only", action="store_true")
    args = ap.parse_args()
    if os.name != "nt":
        ap.error("Windows and desktop Microsoft Excel are required")
    source = args.input.resolve(strict=True)
    output = (args.output or default_output(source)).resolve()
    if source.suffix.lower() not in (".xls", ".xlsx"):
        ap.error("Only .xls and .xlsx inputs are supported")
    if output.suffix.lower() != ".xlsx" or output == source:
        ap.error("Output must be a separate .xlsx")
    if output.exists() and not args.check_only:
        ap.error("Output already exists; specify a different --output (no automatic overwrite)")
    original = hashlib.sha256(source.read_bytes()).digest()
    with tempfile.TemporaryDirectory(prefix="forestry-field-") as tmp:
        raw = Path(tmp) / "source.json"
        excel({"mode": "read", "input": str(source), "sheet": args.sheet, "result": str(raw)}, tmp)
        data = json.loads(raw.read_text(encoding="utf-8-sig"))
        rows, warnings = prepare(data["headers"], data["rows"], area_field=args.area_field, id_field=args.id_field, city=args.city, county=args.county, township=args.township)
        if not args.check_only:
            staging = Path(tmp) / "validated.xlsx"
            excel({"mode": "write", "template": str(ROOT / "assets/A.1.xlsx"), "output": str(staging), "rows": rows}, tmp)
            if hashlib.sha256(source.read_bytes()).digest() != original:
                raise RuntimeError("Source changed during processing; output not published")
            # Exclusive creation prevents racing an existing user workbook.
            with output.open("xb") as dest:
                dest.write(staging.read_bytes())
        print(json.dumps({"records": len(rows), "area_ha": str(sum((Decimal(str(r[7])) for r in rows), Decimal(0))), "warnings": warnings, "output": None if args.check_only else str(output)}, ensure_ascii=False))

if __name__ == "__main__":
    main()
