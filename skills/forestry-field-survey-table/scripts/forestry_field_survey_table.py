"""Prepare independent source-factor and field-survey sheets using desktop Excel."""
import argparse
from decimal import Decimal, InvalidOperation
import hashlib
import json
import os
from pathlib import Path
import subprocess
import tempfile

ROOT = Path(__file__).resolve().parents[1]
# Shipped local data is the source of truth; never fetch personal notes at runtime.
DICTIONARY = json.loads((ROOT / "references/code-dictionary.json").read_text(encoding="utf-8"))
if DICTIONARY.get('schema_version') != 1:
    raise ValueError('Unsupported bundled dictionary schema')
CODE_FIELDS = DICTIONARY['fields']
MAPS = {field: {key: spec.get('display', {}).get(key, name) for key, name in spec['codes'].items()}
        for field, spec in CODE_FIELDS.items() if spec['usage'] == 'automatic'}
LAND_TYPES = dict(DICTIONARY['land_type_names'])
for key, land_type in CODE_FIELDS['LIN_ZHONG']['land_types'].items():
    LAND_TYPES[CODE_FIELDS['LIN_ZHONG']['codes'][key]] = land_type
    LAND_TYPES[MAPS['LIN_ZHONG'][key]] = land_type
SHRUB_NAMES = {"灌木林地", "一般灌木林地", "特殊灌木林地", "国家特别规定灌木林地"}
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

def prepare(headers, source_rows, *, area_field="面积", id_field="小班编", city=None, county=None, township=None, soil_dictionary=None):
    if soil_dictionary not in (None, 'revision-reference'):
        raise ValueError('Unsupported soil dictionary')
    maps = dict(MAPS)
    if soil_dictionary:
        maps['TU_RANG_LX'] = CODE_FIELDS['TU_RANG_LX']['codes']
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
            elif field in maps:
                raw = code(v)
                row[col] = maps[field].get(raw, raw)
                if raw not in maps[field] and raw not in maps[field].values() and raw not in CODE_FIELDS[field]['codes'].values() and not (field == "LIN_ZHONG" and raw in LAND_TYPES):
                    warnings.add(f"Untranslated {field}: {raw}")
            else:
                row[col] = code(v)
        soil = code(r.get('TU_RANG_LX'))
        if not soil_dictionary and soil and soil not in CODE_FIELDS['TU_RANG_LX']['codes'].values():
            if soil in CODE_FIELDS['TU_RANG_LX']['codes']:
                warnings.add(f"Known TU_RANG_LX requires source confirmation: {soil} (revision-reference)")
            else:
                warnings.add(f"Untranslated TU_RANG_LX: {soil}")
        authority = code(r.get('SHI_QUAN_D'))
        if authority and authority not in MAPS['SHI_QUAN_D'] and authority not in MAPS['SHI_QUAN_D'].values():
            warnings.add(f"Untranslated SHI_QUAN_D: {authority}")
        row[6], row[7] = ident, area
        for col, value in enumerate((city, county, township)):
            if value:
                row[col] = value
        tower = code(r.get("塔号"))
        row[32] = tower  # compact enough for the fixed Remarks column
        if row[8] in SHRUB_NAMES:
            row[8] = "一般灌木林地"
        row[9] = "苗圃地" if row[8] == "苗圃地" else LAND_TYPES.get(row[13])
        if row[9] is None:
            warnings.add(f"小班{ident}：林种未匹配林地类型，保留林种原值，林地类型留空")
        if row[21] is not None:
            warnings.add("每公顷蓄积为源库参考值；不据此反算或覆盖核定小班蓄积")
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


def dictionary_coverage(headers, rows, soil_dictionary=None):
    """Report known-but-gated values separately from absent dictionary entries."""
    result = {}
    for index, field in enumerate(map(clean, headers)):
        if field not in CODE_FIELDS:
            continue
        spec = CODE_FIELDS[field]
        values = {code(row[index]) for row in rows if index < len(row)} - {None, ''}
        known = set(spec['codes']) | set(spec['codes'].values()) | set(spec.get('display', {}).values())
        if field == 'LIN_ZHONG':
            known |= set(LAND_TYPES)
        gated = values & set(spec['codes']) if spec['usage'] == 'confirmed_soil' and not soil_dictionary else set()
        result[field] = {'recognized_values': sorted(values & known),
                         'requires_source_confirmation': sorted(gated),
                         'unknown_values': sorted(values - known),
                         'reference_only': spec['usage'] == 'reference_only'}
    return result


def workflow_rows(headers, source_rows, rows, use_nature=None):
    """Keep the source snapshot separate from blank on-site observation columns."""
    by_id = {code(r[headers.index('_id')]): dict(zip(headers, r)) for r in source_rows
             if code(r[headers.index('_id')])}
    original, field = [], []
    for old in rows:
        raw = by_id[old[6]]
        authority = code(raw.get('SHI_QUAN_D'))
        authority = MAPS['SHI_QUAN_D'].get(authority, authority)
        # A:AI; no project-specific old parcel number, no inferred stock/count.
        r = old[:12] + [authority] + old[12:18] + [None] + old[18:]
        r[33] = use_nature
        original.append(r)
        # Second of each paired column is deliberately blank for field entry.
        indices = [6,7,8,None,9,10,11,12,13,14,15,16,17,None,18,19,
                   20,None,21,None,22,None,23,24,25,None,26,None,27,None,
                   28,None,29,None,30,None,34]
        field.append([None if i is None else r[i] for i in indices])
    return [{'name': '原始因子表', 'rows': original}, {'name': '外业调查表', 'rows': field}]

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
    ap.add_argument("--project-name", help="Confirmed full project name for worksheet titles")
    ap.add_argument("--use-nature", choices=['临时', '永久'], help='Confirmed use nature; blank if omitted')
    ap.add_argument("--soil-dictionary", choices=["revision-reference"], help="Enable bundled soil names only after source-version confirmation")
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
        rows, warnings = prepare(data["headers"], data["rows"], area_field=args.area_field, id_field=args.id_field, city=args.city, county=args.county, township=args.township, soil_dictionary=args.soil_dictionary)
        headers = [clean(h) for h in data['headers']]
        headers[headers.index(args.id_field)] = '_id'
        sheets = workflow_rows(headers, data['rows'], rows, args.use_nature)
        project_name = (args.project_name or '').strip()
        for sheet in sheets:
            sheet['title'] = ' '.join(filter(None, [project_name or '建设项目',
                (args.use_nature or '') + '使用林地', sheet['name']]))
        if not args.check_only:
            staging = Path(tmp) / "validated.xlsx"
            excel({"mode": "write", "template": str(ROOT / "assets/A.1.xlsx"), "output": str(staging), "sheets": sheets}, tmp)
            if hashlib.sha256(source.read_bytes()).digest() != original:
                raise RuntimeError("Source changed during processing; output not published")
            # Exclusive creation prevents racing an existing user workbook.
            with output.open("xb") as dest:
                dest.write(staging.read_bytes())
        print(json.dumps({"dictionary_version": DICTIONARY["dictionary_version"], "dictionary_coverage": dictionary_coverage(data['headers'], data['rows'], args.soil_dictionary), "records": len(rows), "area_ha": str(sum((Decimal(str(r[7])) for r in rows), Decimal(0))), "warnings": warnings, "output": None if args.check_only else str(output)}, ensure_ascii=False))

if __name__ == "__main__":
    main()
