"""Stage forestry ownership certificates from explicitly mapped Excel rows; stdlib only."""
import argparse
from collections import Counter
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ET.register_namespace('w', W)


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def write_json(path, value):
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')


def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def require_text(obj, key):
    value = obj.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f'Missing confirmed text: {key}')
    return value.strip()


def decimal(value, field):
    if value is None or isinstance(value, bool) or str(value).strip() == '':
        raise ValueError(f'Missing value: {field}')
    try:
        number = Decimal(str(value))
    except InvalidOperation as exc:
        raise ValueError(f'Invalid number: {field}') from exc
    if not number.is_finite() or number < 0:
        raise ValueError(f'Invalid nonnegative number: {field}')
    return number


def spaced(text):
    return re.sub(r'(?<=[\u3400-\u9fff])(?=[A-Za-z0-9])|(?<=[A-Za-z0-9])(?=[\u3400-\u9fff])', ' ', text)


def compress_ids(ids):
    """Compress only ordinary positive integer IDs; preserve lexical IDs verbatim."""
    if not all(re.fullmatch(r'[1-9][0-9]*', i) for i in ids):
        return '、'.join(ids)
    numbers = sorted(int(i) for i in ids)
    ranges = []
    start = previous = numbers[0]
    for number in numbers[1:]:
        if number != previous + 1:
            ranges.append(str(start) if start == previous else f'{start}-{previous}')
            start = number
        previous = number
    ranges.append(str(start) if start == previous else f'{start}-{previous}')
    return '、'.join(ranges)


def row_list(value, field):
    if not isinstance(value, list) or not value or any(type(x) is not int or x < 1 for x in value):
        raise ValueError(f'Expected explicit Excel row numbers: {field}')
    if len(set(value)) != len(value):
        raise ValueError(f'Duplicate row: {field}')
    return value


def infer_project_dir(source_path):
    matches = [p.parent for p in Path(source_path).resolve().parents if p.name == '05_报告表格与图件']
    if len(matches) != 1:
        raise ValueError('Cannot uniquely locate project root from 05_报告表格与图件; specify project_dir')
    return matches[0]


def default_target(c):
    return Path(c['project_dir']).resolve() / '06_公示查验与报批' / c['use_nature'] / '林木权属证明' / f'林木权属证明（{c["use_nature"]}）.docx'


def validate_config(c):
    source = c['source']
    source.setdefault('sheet', 'A.1')
    if 'project_dir' not in c:
        c['project_dir'] = str(infer_project_dir(require_text(source, 'path')))
    for key in ('project_name', 'project_name_basis', 'applicant', 'applicant_basis', 'project_dir'):
        require_text(c, key)
    if c.get('use_nature') not in ('永久', '临时'):
        raise ValueError('use_nature must be 永久 or 临时')
    if source.get('confirmed') is not True:
        raise ValueError('Source fields, units and scope require confirmation')
    for key in ('path', 'sheet', 'area_basis', 'stock_basis'):
        require_text(source, key)
    if Path(source['path']).suffix.lower() not in ('.xls', '.xlsx'):
        raise ValueError('Source must be XLS or XLSX')
    rows = row_list(source['data_rows'], 'data_rows')
    columns = source.get('columns', {})
    if not isinstance(columns, dict):
        raise ValueError(f'Invalid field mapping in sheet {source["sheet"]}')
    missing = {'id', 'area', 'stock'} - columns.keys()
    if missing:
        raise ValueError(f'Missing required field mapping in sheet {source["sheet"]}: {", ".join(sorted(missing))}')
    if set(columns) - {'id', 'area', 'stock', 'trees'}:
        raise ValueError('Only id, area, stock and optional trees mappings are supported')
    if any(not isinstance(v, str) or not re.fullmatch(r'[A-Z]{1,3}', v) for v in columns.values()):
        raise ValueError('Use uppercase Excel column letters')
    if any(sum((ord(ch) - 64) * 26 ** i for i, ch in enumerate(reversed(v))) > 16384
           for v in columns.values()):
        raise ValueError('Excel columns cannot exceed XFD')
    if len(set(columns.values())) != len(columns):
        raise ValueError('Source columns must be distinct')
    if not isinstance(c.get('groups'), list) or not c['groups']:
        raise ValueError('At least one confirmed ownership group is required')
    assigned = []
    keys = []
    for g in c['groups']:
        for key in ('city', 'county', 'village', 'owner', 'ownership_basis'):
            require_text(g, key)
        if g.get('kind') not in ('集体', '国有'):
            raise ValueError('Only confirmed collective and state ownership are supported')
        if g['kind'] == '集体':
            require_text(g, 'township')
        if g.get('ownership_confirmed') is not True or g.get('undisputed_confirmed') is not True:
            raise ValueError('Ownership and no-dispute statement must be explicitly confirmed')
        if g.get('trees_mode') == 'include':
            if 'trees' not in columns:
                raise ValueError('Total tree count column required')
            require_text(source, 'trees_basis')
        elif g.get('trees_mode') == 'omit':
            require_text(g, 'trees_omit_basis')
        else:
            raise ValueError('trees_mode must explicitly be include or omit')
        if not isinstance(g.get('note', ''), str):
            raise ValueError('note must be confirmed plain text')
        assigned.extend(row_list(g['rows'], 'group.rows'))
        keys.append(tuple(g.get(k, '') for k in ('city', 'county', 'township', 'village', 'owner')))
    if Counter(assigned) != Counter(rows):
        raise ValueError('Group rows must cover every data row exactly once')
    if len(set(keys)) != len(keys):
        raise ValueError('Duplicate ownership groups: combine rows for the same owner and location')


def prepare(c, records):
    validate_config(c)
    by_row = {r['row']: r for r in records}
    if len(by_row) != len(records) or set(by_row) != set(c['source']['data_rows']):
        raise ValueError('Extracted rows differ from confirmed scope')
    ids = []
    for r in records:
        identifier = '' if r.get('id') is None else str(r['id']).strip()
        if not identifier or '#' in identifier or '\n' in identifier or '\r' in identifier:
            raise ValueError(f'Invalid parcel ID at row {r["row"]}')
        r['id'] = identifier
        ids.append(identifier)
    if len(set(ids)) != len(ids):
        raise ValueError('Duplicate parcel IDs within this project and use nature')
    certificates = []
    for g in c['groups']:
        selected = [by_row[row] for row in g['rows']]
        area = sum((decimal(r.get('area'), f'area row {r["row"]}') for r in selected), Decimal(0))
        if any(decimal(r.get('area'), 'area') <= 0 for r in selected):
            raise ValueError('Occupied parcel area must be positive')
        stock = sum((decimal(r.get('stock'), f'stock row {r["row"]}') for r in selected), Decimal(0))
        trees = None
        if g['trees_mode'] == 'include':
            counts = [decimal(r.get('trees'), f'trees row {r["row"]}') for r in selected]
            if any(n != n.to_integral_value() for n in counts):
                raise ValueError('Total tree count must be an integer, never silently rounded')
            trees = str(int(sum(counts)))
        summary = {
            **g, 'ids': [r['id'] for r in selected], 'area_raw': str(area), 'stock_raw': str(stock),
            'area': str(area.quantize(Decimal('.0001'), rounding=ROUND_HALF_UP)),
            'stock': str(stock.quantize(Decimal('1'), rounding=ROUND_HALF_UP)), 'trees': trees,
        }
        location = ''.join(g.get(k, '') for k in ('city', 'county', 'township', 'village'))
        body = (f'因{spaced(c["applicant"])}实施“{spaced(c["project_name"])}”，涉及{c["use_nature"]}使用林地，'
                f'现就位于{location}的使用林地范围出具林木权属证明，涉及地块编号 {compress_ids(summary["ids"])} 号地块。'
                f'涉及地块面积：{summary["area"]} 公顷，')
        if trees is not None:
            body += f'林木株数：{trees} 株，'
        body += f'蓄积：{summary["stock"]} 立方米。'
        if g.get('note'):
            body += spaced(g['note'])
        body += f'上述地块内的林木均为{g["owner"]}所有。林木权属无争议。'
        summary['body'] = spaced(body)
        certificates.append(summary)
    return certificates


def el(tag, attrs=None, text=None):
    node = ET.Element(f'{{{W}}}{tag}', {f'{{{W}}}{k}': str(v) for k, v in (attrs or {}).items()})
    node.text = text
    return node


def paragraph(text, size=16, bold=False, center=False, before=0, after=0, page=False, indent=0, line=624):
    p = el('p')
    prop = ET.SubElement(p, f'{{{W}}}pPr')
    if page:
        prop.append(el('pageBreakBefore'))
    prop.append(el('spacing', {'before': before * 20, 'after': after * 20, 'line': line, 'lineRule': 'exact'}))
    prop.append(el('ind', {'firstLine': indent * 20}))
    prop.append(el('jc', {'val': 'center' if center else 'left'}))
    run = ET.SubElement(p, f'{{{W}}}r')
    rp = ET.SubElement(run, f'{{{W}}}rPr')
    rp.append(el('rFonts', {'ascii': 'Times New Roman', 'hAnsi': 'Times New Roman', 'eastAsia': '宋体'}))
    rp.append(el('sz', {'val': size * 2}))
    if bold:
        rp.append(el('b'))
    t = el('t', text=spaced(text))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    run.append(t)
    return p


def signing_table(kind):
    table = el('tbl')
    props = el('tblPr')
    props.append(el('tblW', {'w': 8306, 'type': 'dxa'}))
    props.append(el('tblLayout', {'type': 'fixed'}))
    borders = el('tblBorders')
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        borders.append(el(side, {'val': 'nil'}))
    props.append(borders)
    margins = el('tblCellMar')
    for side, width in (('top', 0), ('bottom', 0), ('left', 0), ('right', 80)):
        margins.append(el(side, {'w': width, 'type': 'dxa'}))
    props.append(margins)
    table.append(props)
    grid = el('tblGrid')
    for _ in range(2):
        grid.append(el('gridCol', {'w': 4153}))
    table.append(grid)
    row = el('tr')
    rp = el('trPr'); rp.append(el('cantSplit')); row.append(rp)
    stamps = ('村委会盖章：', '乡镇政府盖章：') if kind == '集体' else ('林场盖章：', '主管单位盖章：')
    for stamp in stamps:
        cell = el('tc'); cp = el('tcPr'); cp.append(el('tcW', {'w': 4153, 'type': 'dxa'})); cell.append(cp)
        for i, text in enumerate(('签字：', stamp, '　　　　年　　月　　日')):
            p = paragraph(text, size=14, line=560)
            if i < 2:
                p.find(f'{{{W}}}pPr').append(el('keepNext'))
            cell.append(p)
        row.append(cell)
    table.append(row)
    return table


def write_docx(path, config, certs):
    document = el('document'); body = el('body'); document.append(body)
    body.append(paragraph('林木权属证明签章对接备注', 18, True, True, after=18))
    body.append(paragraph(spaced(config['project_name']), 14, True, True, after=12))
    body.append(paragraph(f'使用性质：{config["use_nature"]}使用林地', 12, center=True, after=18))
    body.append(paragraph(f'本册收录 {len(certs)} 份林木权属证明，正文按下列对接单位的顺序排列。', 12, after=12, line=360))
    groups = []
    locations = Counter(tuple(g.get(k, '') for k in ('city', 'county', 'township', 'village')) for g in certs)
    for g in certs:
        key = tuple(g.get(k, '') for k in ('city', 'county', 'township', 'kind'))
        village = g['village']
        if locations[tuple(g.get(k, '') for k in ('city', 'county', 'township', 'village'))] > 1:
            village += f'（{g["owner"]}）'
        if groups and groups[-1][0] == key:
            groups[-1][1].append(village)
        else:
            groups.append((key, [village]))
    for key, villages in groups:
        city, county, town, kind = key
        place = ' '.join(s for s in (city, county, town) if s)
        suffix = f'（{config["use_nature"]}' + ('，国有）' if kind == '国有' else '）')
        body.append(paragraph(f'{place}：' + '、'.join(villages) + suffix, 12, after=10, line=360))
    body.append(paragraph('签章对接事项', 13, True, after=10, line=360))
    collective = any(g['kind'] == '集体' for g in certs)
    state = any(g['kind'] == '国有' for g in certs)
    lines = []
    if collective:
        lines.append('第一阶段：分别对接各村，完成负责人签字、村委会盖章及日期填写。')
    if state:
        lines.append(('国有林场同步' if collective else '第一阶段：对接国有林场，') + '办理负责人签字、林场盖章及日期填写。')
    lines.append('第二阶段：按项目统一收齐签章材料。')
    if collective:
        lines.append('村级材料按所属乡镇归集，集中送相应乡镇人民政府办理签字、盖章及日期填写。')
    if state:
        lines.append('林场材料一并集中办理主管单位签字、盖章及日期填写。')
    lines.append('本页仅供签章对接使用，正式报送时不装入证明材料。')
    for text in lines:
        body.append(paragraph(text, 12, after=10, line=360))
    for g in certs:
        title = paragraph('林木权属证明', 22, center=True, before=24, after=24, page=True)
        title.find(f'{{{W}}}pPr').append(el('keepNext')); body.append(title)
        body.append(paragraph(g['body'], indent=32))
        body.append(paragraph('特此证明。', indent=32, after=54))
        body.append(signing_table(g['kind']))
    section = el('sectPr'); section.append(el('pgSz', {'w': 11906, 'h': 16838}))
    section.append(el('pgMar', {'top': 1440, 'bottom': 1440, 'left': 1800, 'right': 1800, 'header': 720, 'footer': 720, 'gutter': 0}))
    body.append(section)
    types = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', types)
        z.writestr('_rels/.rels', rels)
        z.writestr('word/document.xml', ET.tostring(document, encoding='utf-8', xml_declaration=True))


def office(job, directory):
    job_path = Path(directory) / f'{job["mode"]}-job.json'
    write_json(job_path, job)
    executable = shutil.which('pwsh') or shutil.which('powershell')
    if not executable:
        raise RuntimeError('Windows PowerShell and desktop Office are required')
    subprocess.run([executable, '-NoProfile', '-NonInteractive', '-File', str(ROOT / 'scripts/office_io.ps1'), '-Job', str(job_path)], check=True)


def publish(stage_path):
    stage = read_json(stage_path)
    for entry in stage['inputs']:
        if digest(entry['path']) != entry['sha256']:
            raise ValueError(f'Input changed: {entry["path"]}')
    candidate = Path(stage['candidate'])
    if digest(candidate) != stage['candidate_sha256']:
        raise ValueError('Candidate changed after validation')
    target = Path(stage['target'])
    if (digest(target) if target.exists() else None) != stage['target_sha256']:
        raise ValueError('Target changed after staging')
    if target.exists() and not stage['replace']:
        raise FileExistsError(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=target.parent, suffix='.docx', delete=False) as f:
        temporary = Path(f.name)
    try:
        shutil.copyfile(candidate, temporary)
        if stage['target_sha256'] is None:
            # Exclusive creation prevents replacing a target created after the preflight.
            with target.open('xb') as out, temporary.open('rb') as src:
                shutil.copyfileobj(src, out)
        else:
            if digest(target) != stage['target_sha256']:
                raise ValueError('Target changed during publication')
            os.replace(temporary, target)
        if digest(target) != stage['candidate_sha256']:
            raise ValueError('Published file differs from candidate')
    finally:
        temporary.unlink(missing_ok=True)
    return target


def stage(config_path, output=None, replace=False):
    if os.name != 'nt':
        raise RuntimeError('Generation requires Windows and desktop Microsoft Excel / Word')
    config_path = Path(config_path).resolve(); c = read_json(config_path); validate_config(c)
    source = Path(c['source']['path']).resolve()
    target = Path(output).resolve() if output else default_target(c)
    if target.name != f'林木权属证明（{c["use_nature"]}）.docx' or target in (source, config_path):
        raise ValueError('Output must use the confirmed certificate filename and differ from inputs')
    if target.exists() and not replace:
        raise FileExistsError(f'Target exists; review manual edits before authorized --replace: {target}')
    inputs = [{'path': str(p), 'sha256': digest(p)} for p in (source, config_path)]
    target_hash = digest(target) if target.exists() else None
    directory = Path(tempfile.mkdtemp(prefix='forestry-ownership-'))
    if target.exists():
        shutil.copy2(target, directory / 'previous.docx')
    records_path = directory / 'records.json'
    office({'mode': 'read', 'source': {**c['source'], 'path': str(source)}, 'result': str(records_path)}, directory)
    records = read_json(records_path)
    certs = prepare(c, records)
    candidate = directory / 'candidate.docx'; write_docx(candidate, c, certs)
    audit = {'project_name': c['project_name'], 'use_nature': c['use_nature'], 'certificates': certs}
    write_json(directory / 'audit.json', audit)
    office({'mode': 'verify', 'candidate': str(candidate), 'pdf': str(directory / 'preview.pdf'), 'bodies': [g['body'] for g in certs], 'count': len(certs)}, directory)
    for entry in inputs:
        if digest(entry['path']) != entry['sha256']:
            raise ValueError('Source or config changed during generation')
    manifest = {'inputs': inputs, 'candidate': str(candidate), 'candidate_sha256': digest(candidate), 'target': str(target), 'target_sha256': target_hash, 'replace': replace}
    write_json(directory / 'stage.json', manifest)
    return directory / 'stage.json'


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument('--config'); modes.add_argument('--publish-stage')
    parser.add_argument('--output'); parser.add_argument('--replace', action='store_true')
    args = parser.parse_args()
    if args.publish_stage and (args.output or args.replace):
        parser.error('--output and --replace apply only to staging')
    try:
        result = publish(args.publish_stage) if args.publish_stage else stage(args.config, args.output, args.replace)
        print(result)
    except (ValueError, KeyError, OSError, RuntimeError, subprocess.CalledProcessError) as exc:
        parser.exit(1, f'{exc}\n')


if __name__ == '__main__':
    main()
