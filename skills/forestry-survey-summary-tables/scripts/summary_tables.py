"""Stage and publish forestry report tables using native Excel; Python stdlib only."""
import argparse
from copy import deepcopy
from decimal import Decimal
import hashlib
import json
import math
import os
from pathlib import Path
import posixpath
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
REPORT_NAMES = [f'B.{i}' for i in range(1, 8)] + ['小班明细']
REPLACED_NAMES = set(REPORT_NAMES) | {str(i) for i in range(1, 8)}
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
TYPES = ['防护林林地', '特用林林地', '用材林林地', '经济林林地', '能源林林地', '苗圃地', '其他林地']
CLASSES = ['乔木林地', '竹林地', '国家特别规定灌木林地', '一般灌木林地', '疏林地', '未成林造林地', '苗圃地', '采伐迹地', '火烧迹地', '其他林地']
FORESTS = ['国家级一级公益林', '国家级二级公益林', '国家级三级公益林', '省级公益林', '市县级公益林', '重点商品林', '一般商品林']
ECO = ['自然保护区', '森林公园', '湿地公园', '风景名胜区']
HEADERS = ['市', '县', '乡', '村', '林班号(村小组)', '林地落界小班号', '小班编号', '面积hm2', '现状地类', '林地类型', '林地权属', '森林类别', '林地保护等级', '林种', '起源', '重点生态区域', '优势树种', '龄组(产期)', '平均树高m', '平均胸径cm', '郁闭度(覆盖度)', '每公顷蓄积m3', '小班蓄积m3', '经济(竹)林株数', '坡向', '坡位', '坡度', '土壤类型', '土层厚度', '建设内容', '工程类别', '使用林地性质', '备注']


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def col(n):
    s = ''
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def normal(s):
    return ''.join(str(s or '').split()).replace('（', '(').replace('）', ')').replace('²', '2').replace('³', '3')


def read_book(path):
    """Read cached cells and XML, without modifying the native workbook."""
    with zipfile.ZipFile(path) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = {r.attrib['Id']: r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            shared = [''.join(t.text or '' for t in si.findall('.//s:t', NS)) for si in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        result = {}
        styles = ET.fromstring(z.read('xl/styles.xml'))
        for sh in wb.find('s:sheets', NS):
            key = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
            part = rels[sh.attrib[key]]
            part = part.lstrip('/') if part.startswith('/') else posixpath.normpath('xl/' + part)
            xml = ET.fromstring(z.read(part))
            cells = {}
            for c in xml.findall('.//s:sheetData/s:row/s:c', NS):
                t = c.attrib.get('t'); v = c.find('s:v', NS); f = c.find('s:f', NS)
                value = v.text if v is not None else None
                if t == 's': value = shared[int(value)]
                elif t == 'inlineStr': value = ''.join(e.text or '' for e in c.findall('.//s:t', NS))
                elif t not in ('str', 'e') and value is not None:
                    value = float(value)
                    if value.is_integer(): value = int(value)
                cells[c.attrib['r']] = {'value': value, 'formula': f.text if f is not None else None, 'type': t, 'style': int(c.attrib.get('s', 0))}
            result[sh.attrib['name']] = {'cells': cells, 'xml': xml, 'styles': styles}
        if any(n.startswith('xl/externalLinks/externalLink') and n.endswith('.xml') for n in z.namelist()):
            raise ValueError('External workbook links are not supported')
        return result


def val(sheet, address):
    return sheet['cells'].get(address, {}).get('value')


def number(v, label, positive=False):
    if isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) or v < 0 or (positive and v == 0):
        raise ValueError(f'Invalid or missing {label}: {v}')
    return v


def text(v):
    return str(v).strip() if v is not None else ''


def prepare(book, cfg):
    for k in ['project_name', 'use_nature', 'forest_category_map', 'ecology', 'fee']:
        if k not in cfg: raise ValueError(f'Missing configuration: {k}')
    if not text(cfg['project_name']) or cfg['use_nature'] not in ('临时', '永久'):
        raise ValueError('Project name and use nature must be confirmed')
    s = book['A.1']
    for i, h in enumerate(HEADERS, 1):
        if normal(val(s, f'{col(i)}3')) != normal(h): raise ValueError(f'Unexpected A.1 header: {col(i)}3')
    last = max(int(''.join(filter(str.isdigit, a))) for a in s['cells'])
    records, seen = [], set()
    for r in range(4, last + 1):
        vs = [val(s, f'{col(c)}{r}') for c in range(1, 34)]
        if all(v in (None, '') for v in vs): continue
        ident = text(vs[6])
        if not ident or ident in seen: raise ValueError(f'Missing/duplicate parcel ID: {ident}')
        seen.add(ident)
        over = cfg.get('parcel_overrides', {}).get(ident, {})
        if set(over) - {'land_class', 'land_type', 'forest_category'}: raise ValueError(f'Unknown override for {ident}')
        lc = over.get('land_class', cfg.get('land_class_map', {}).get(text(vs[8]), text(vs[8])))
        lt = over.get('land_type', cfg.get('land_type_map', {}).get(text(vs[9]), text(vs[9])))
        if lt == '薪炭林林地': lt = '能源林林地'
        fc = over.get('forest_category', cfg['forest_category_map'].get(text(vs[11])))
        level = {'1': 'Ⅰ', '2': 'Ⅱ', '3': 'Ⅲ', '4': 'Ⅳ', 'I': 'Ⅰ', 'II': 'Ⅱ', 'III': 'Ⅲ', 'IV': 'Ⅳ'}.get(text(vs[12]), text(vs[12]))
        if lc not in CLASSES or lt not in TYPES or fc not in FORESTS or level not in ['Ⅰ', 'Ⅱ', 'Ⅲ', 'Ⅳ']:
            raise ValueError(f'Unresolved classification for parcel {ident}: {lc}, {lt}, {fc}, {level}')
        if any(not text(v) for v in vs[:4]) or vs[10] not in ('国有', '集体', '个人'):
            raise ValueError(f'Incomplete administration/ownership for parcel {ident}')
        if text(vs[31]) not in ('', cfg['use_nature']): raise ValueError(f'Use nature conflict: {ident}')
        area = number(vs[7], f'area ({ident})', True)
        volume = number(vs[22], f'tree volume ({ident})') if lc == '乔木林地' else 0
        records.append(dict(id=ident, source=r, values=vs, admin=tuple(map(text, vs[:4])), own=vs[10], land_class=lc, land_type=lt, forest_category=fc, level=level, area=area, volume=volume))
    if not records: raise ValueError('No A.1 records')
    if set(cfg.get('parcel_overrides', {})) - seen: raise ValueError('Unknown parcel override ID')
    eco = cfg['ecology']
    if eco.get('mode') == 'none':
        if eco.get('confirmed') is not True: raise ValueError('Ecological non-involvement needs confirmation')
        if any(text(r['values'][15]) not in ('', '其他区域', '不涉及', '无') for r in records):
            raise ValueError('Ecology none conflicts with A.1')
    elif eco.get('mode') == 'records':
        covered, keys = set(), set()
        if not eco.get('records'): raise ValueError('Empty ecology records')
        for row in eco['records']:
            ids = row.get('parcel_ids', [])
            key = (row.get('kind'), row.get('name'))
            if not text(row.get('name')) or row.get('kind') not in ECO or not text(row.get('level')) or key in keys:
                raise ValueError('Invalid/duplicate ecological region')
            if not ids or len(set(ids)) != len(ids) or set(ids) - seen: raise ValueError('Invalid ecological parcel IDs')
            covered.update(ids);keys.add(key)
        if any(r['id'] not in covered and text(r['values'][15]) not in ('', '其他区域', '不涉及', '无') for r in records):
            raise ValueError('Unmapped ecological involvement')
    else: raise ValueError('Unsupported ecology mode')
    fee = cfg['fee']
    if fee.get('mode') not in ('reference', 'charge', 'not_applicable') or not text(fee.get('basis')):
        raise ValueError('Fee mode and verified basis are required')
    if cfg['use_nature'] == '临时' and fee['mode'] == 'charge': raise ValueError('Temporary use is not an actual fee charge')
    if fee['mode'] != 'not_applicable':
        if not text(fee.get('context')) or not text(fee.get('note')): raise ValueError('Fee context/note is required')
        number(fee.get('default_multiplier'), 'default multiplier')
        if set(fee.get('parcel_multipliers', {})) - seen: raise ValueError('Unknown fee multiplier ID')
        matches = {}
        for cat in fee.get('categories', []):
            if not text(cat.get('name')) or not cat.get('land_classes'): raise ValueError('Invalid fee category')
            number(cat.get('rate'), 'rate')
            for lc in cat['land_classes']:
                if lc not in CLASSES or lc in matches: raise ValueError('Unknown/overlapping fee land class')
                matches[lc] = cat
        if not 1 <= len(fee.get('categories', [])) <= 10: raise ValueError('Fee table supports 1-10 verified categories')
        for r in records:
            if r['land_class'] not in matches: raise ValueError(f'Fee classification missing: {r["id"]}')
            multiplier = number(fee.get('parcel_multipliers', {}).get(r['id'], fee['default_multiplier']), 'parcel multiplier')
            r['rate'] = matches[r['land_class']]['rate'] * multiplier
    return records


def groups(records):
    result = [()]
    def walk(prefix):
        if len(prefix) == 4: return
        children = list(dict.fromkeys(r['admin'][:len(prefix)+1] for r in records if r['admin'][:len(prefix)] == prefix))
        for child in children:
            result.append(child);walk(child)
    walk(())
    return result


def make_sheet(n, head, cols):
    return dict(name=str(n), head=head, cols=cols, cells={}, merges=[], formats=[], heights={}, checks={})


def setcell(s, address, value):
    s['cells'][address] = value


def merge(s, address, value):
    s['merges'].append(address)
    setcell(s, address.split(':')[0], value)


def aggregate(records, column):
    # Compress adjacent detail rows; summaries never include other subtotal rows.
    rows = sorted(r['detail'] for r in records)
    terms = []
    for row in rows:
        if terms and row == terms[-1][1] + 1: terms[-1][1] = row
        else: terms.append([row, row])
    refs = [f"'小班明细'!{column}{a}" + (f':{column}{b}' if a != b else '') for a, b in terms]
    if not refs: return '=0'
    # Each SUM call stays below Excel's 255-argument limit.
    formula = '=' + '+'.join('SUM(' + ','.join(refs[i:i+200]) + ')' for i in range(0, len(refs), 200))
    if len(formula) > 8000: raise ValueError('Too many disjoint rows for a safe formula; split the workbook scope')
    return formula


def expected(records, metric):
    return float(sum((Decimal(str(r[metric])) for r in records), Decimal(0)))


def put_sum(s, address, records, metric='area'):
    setcell(s, address, aggregate(records, 'H' if metric == 'area' else 'W'))
    s['checks'][address] = expected(records, metric)


def labels(s, gs, own, start, shift):
    # One total row followed by ownership rows at each actual administrative node.
    size = len(own)
    for i, g in enumerate(gs):
        r = start + i * size
        label = ['总计', '合计', '小计', '计', g[-1] if g else ''][len(g)]
        c = shift + min(len(g), 3)
        merge(s, f'{col(c)}{r}:{col(shift+3)}{r+size-1}', label)
        for j, o in enumerate(own): setcell(s, f'{col(shift+4)}{r+j}', o)
        if g and len(g) < 4:
            end = i
            while end+1 < len(gs) and gs[end+1][:len(g)] == g: end += 1
            merge(s, f'{col(shift+len(g)-1)}{r}:{col(shift+len(g)-1)}{start+(end+1)*size-1}', g[-1])


def cross_table(records, cfg):
    """B.5: county/forest category by land class and ownership, area only."""
    classes = CLASSES[:5] + [c for c in CLASSES[5:] if any(r['land_class'] == c for r in records)]
    owners = ['国有', '集体'] + (['个人'] if any(r['own'] == '个人' for r in records) else [])
    nc = 3 + (1 + len(classes)) * len(owners)
    s = make_sheet('B.5', 6, nc)
    merge(s, f'A1:{col(nc)}2', f'表B.5 {cfg["project_name"]}\n{cfg["use_nature"]}使用林地分森林类别按地类面积统计表')
    merge(s, f'A3:{col(nc)}3', '单位：公顷')
    merge(s, 'A4:A5', '县、林业局');merge(s, 'B4:C5', '森林类别')
    setcell(s, 'A6', 1);merge(s, 'B6:C6', 2)
    for k, lc in enumerate(['合计'] + classes):
        c = 4 + k * len(owners)
        merge(s, f'{col(c)}4:{col(c+len(owners)-1)}4', lc)
        for j, owner in enumerate(owners):
            setcell(s, f'{col(c+j)}5', owner);setcell(s, f'{col(c+j)}6', c+j-1)
    counties = list(dict.fromkeys(r['admin'][:2] for r in records))
    categories = [('小计', FORESTS[:5]), ('国家级公益林地', FORESTS[:3]),
                  ('省级公益林地', FORESTS[3:4]), ('市县级公益林地', FORESTS[4:5]),
                  ('小计', FORESTS[5:]), ('重点商品林地', FORESTS[5:6]), ('一般商品林地', FORESTS[6:])]
    for i, county in enumerate([None] + counties):
        start = 7 + 7*i
        # Include city only when the same county name occurs in multiple cities.
        name = '合计' if county is None else (''.join(county) if sum(c[1] == county[1] for c in counties) > 1 else county[1])
        merge(s, f'A{start}:A{start+6}', name)
        merge(s, f'B{start}:B{start+3}', '公益林地');merge(s, f'B{start+4}:B{start+6}', '商品林地')
        for j, (label, forests) in enumerate(categories):
            rr = start+j;setcell(s, f'C{rr}', label)
            base = [r for r in records if (county is None or r['admin'][:2] == county) and r['forest_category'] in forests]
            for k, lc in enumerate([None] + classes):
                for h, owner in enumerate(owners):
                    bucket = [r for r in base if r['own'] == owner and (lc is None or r['land_class'] == lc)]
                    put_sum(s, f'{col(4+k*len(owners)+h)}{rr}', bucket)
    s['last'] = 6+7*(1+len(counties))
    s['formats'] = [(f'A6:{col(nc)}6', '0'), (f'D7:{col(nc)}{s["last"]}', '0.0000;-0.0000;;@')]
    return s


def build(records, cfg):
    gs = groups(records)
    # Same hierarchical nodes for the report detail, with leaf parcel rows.
    detail = make_sheet('小班明细', 3, 28)
    merge(detail, 'A1:AB1', f'{cfg["project_name"]}\n{cfg["use_nature"]}使用林地小班因子调查表（小班明细）')
    row = 4
    nodes = []
    for g in gs:
        items = [x for x in records if x['admin'][:len(g)] == g]
        begin = row
        depth = len(g)
        c = min(depth+1, 5)
        merge(detail, f'{col(c)}{row}:G{row}', ['省总计', '市合计', '县合计', '镇合计', '小计'][depth])
        row += 1
        if depth == 4:
            for rec in items:
                rec['detail'] = row
                for dc in range(5, 29):
                    sc = dc if dc <= 24 else dc+5
                    ref = f"'A.1'!{col(sc)}{rec['source']}"
                    value = f'=IF({ref}="","",{ref})'
                    if dc == 9: value = rec['land_class']
                    if dc == 10: value = rec['land_type']
                    if dc == 12: value = rec['forest_category']
                    if dc == 13: value = rec['level']
                    if dc == 16:
                        value = '不涉及' if cfg['ecology']['mode']=='none' else '、'.join(x['name'] for x in cfg['ecology']['records'] if rec['id'] in x['parcel_ids']) or '不涉及'
                    if dc == 23 and rec['land_class'] != '乔木林地': value = None
                    if dc == 27: value = cfg['use_nature']
                    setcell(detail, f'{col(dc)}{row}', value)
                row += 1
        nodes.append((g, begin, items))
    for i, (g, begin, items) in enumerate(nodes):
        if g:
            end = row-1
            for gg, rr, _ in nodes[i+1:]:
                if gg[:len(g)] != g: end = rr-1;break
            merge(detail, f'{col(len(g))}{begin}:{col(len(g))}{end}', g[-1])
        for metric, c in [('area','H'), ('volume','W')]: put_sum(detail, f'{c}{begin}', items, metric)
        detail['heights'][str(begin)] = 18
    detail['last'] = row-1
    detail['formats'] = [('H4:H'+str(row-1), '0.0000;-0.0000;;@'), ('S4:T'+str(row-1), '0.0;-0.0;;@'), ('U4:V'+str(row-1), '0.##;-0.##;;@'), ('W4:W'+str(row-1), '0.0;-0.0;;@'), ('X4:X'+str(row-1), '0;-0;;@')]
    sheets = []
    own = ['计'] + [o for o in ['国有', '集体', '个人'] if any(r['own']==o for r in records)]
    size = len(gs)*len(own)
    titles = ['按林地类型面积蓄积统计表', '按现状地类面积蓄积统计表', '按森林类别面积蓄积统计表', '按林地保护等级面积蓄积统计表']
    for n in range(1,5):
        head, nc = (7,17) if n==3 else (5,14 if n==1 else 17 if n==2 else 11)
        s = make_sheet(f'B.{n}', head, nc)
        setcell(s,'A1',f'表B.{n} {cfg["project_name"]}\n{cfg["use_nature"]}使用林地{titles[n-1]}')
        for metric, offset in [('area',0), ('volume',size)]:
            start = head+1+offset
            labels(s,gs,own,start,2)
            merge(s,f'A{start}:A{start+size-1}', '面积' if metric=='area' else '蓄积')
            for i,g in enumerate(gs):
                for j,o in enumerate(own):
                    rownum = start+i*len(own)+j
                    base = [r for r in records if r['admin'][:len(g)]==g and (o=='计' or r['own']==o)]
                    if n==1: bins = [base]+[[r for r in base if r['land_type']==t] for t in TYPES]
                    elif n==2: bins = [base]+[[r for r in base if r['land_class']==t] for t in CLASSES]
                    elif n==4: bins = [base]+[[r for r in base if r['level']==t] for t in ['Ⅰ','Ⅱ','Ⅲ','Ⅳ']]
                    else:
                        tests = [FORESTS,FORESTS[:5],FORESTS[:3],FORESTS[:1],FORESTS[1:2],FORESTS[2:3],FORESTS[3:4],FORESTS[4:5],FORESTS[5:],FORESTS[5:6],FORESTS[6:]]
                        bins = [[r for r in base if r['forest_category'] in ts] for ts in tests]
                    for c, bucket in enumerate(bins,7): put_sum(s,f'{col(c)}{rownum}',bucket,metric)
            s['formats'].append((f'G{start}:{col(nc)}{start+size-1}', '0.0000;-0.0000;;@' if metric=='area' else '0.0;-0.0;;@'))
        s['last'] = head+2*size
        sheets.append(s)
    sheets.append(cross_table(records, cfg))
    s = make_sheet('B.6',5,11)
    setcell(s,'A1',f'表B.6 {cfg["project_name"]}\n{cfg["use_nature"]}使用重点生态区域林地面积统计表')
    if cfg['ecology']['mode']=='none': setcell(s,'A6','不涉及');s['last']=6
    else:
        for rownum,reg in enumerate(cfg['ecology']['records'],6):
            items=[r for r in records if r['id'] in reg['parcel_ids']]
            setcell(s,f'A{rownum}',reg['name']);put_sum(s,f'B{rownum}',items)
            c=3+2*ECO.index(reg['kind']);setcell(s,f'{col(c)}{rownum}',reg['level']);put_sum(s,f'{col(c+1)}{rownum}',items)
        s['last']=5+len(cfg['ecology']['records'])
    s['formats']=[(f'B6:J{s["last"]}', '0.0000;-0.0000;;@')];sheets.append(s)
    fee=cfg['fee'];cats=fee.get('categories',[]) if fee['mode']!='not_applicable' else []
    fee_header=read_book(ROOT/'assets/report-headers.xlsx')['B.7']
    nc=max(10,7+len(cats)*3);s=make_sheet('B.7',7,nc)
    merge(s,f'A1:{col(nc)}2',f'表B.7 {cfg["project_name"]}\n{cfg["use_nature"]}使用林地森林植被恢复费'+('参考测算表' if fee['mode']=='reference' else '测算表'))
    merge(s,f'A3:{col(nc)}3','单位：面积—公顷；折算单价—元/平方米；金额—万元')
    if not cats:
        merge(s,f'A4:{col(nc)}7','不适用：'+fee['basis']);s['last']=7
    else:
        for c in range(1,6): merge(s,f'{col(c)}4:{col(c)}7',val(fee_header,f'{col(c)}4'))
        merge(s,'F4:G6',val(fee_header,'F4'));setcell(s,'F7',val(fee_header,'F7'));setcell(s,'G7',val(fee_header,'G7'))
        for i,cat in enumerate(cats):
            c=8+i*3;merge(s,f'{col(c)}4:{col(c+2)}6',cat['name'])
            for j in range(3):setcell(s,f'{col(c+j)}7',val(fee_header,f'{col(8+j)}7'))
        labels(s,gs,own,8,1)
        for i,g in enumerate(gs):
            for j,o in enumerate(own):
                rr=8+i*len(own)+j
                items=[r for r in records if r['admin'][:len(g)]==g and (o=='计' or r['own']==o)]
                put_sum(s,f'F{rr}',items)
                setcell(s,f'G{rr}','=SUM('+','.join(f'{col(10+k*3)}{rr}' for k in range(len(cats)))+')')
                s['checks'][f'G{rr}']=sum(r['area']*r['rate'] for r in items)
                for k,cat in enumerate(cats):
                    c=8+k*3;subset=[r for r in items if r['land_class'] in cat['land_classes']]
                    put_sum(s,f'{col(c)}{rr}',subset)
                    byrate={rate:[r for r in subset if r['rate']==rate] for rate in {r['rate'] for r in subset}}
                    formula='='+ '+'.join(f'({aggregate(rs,"H")[1:]})*{rate}' for rate,rs in sorted(byrate.items())) if subset else '=0'
                    setcell(s,f'{col(c+2)}{rr}',formula)
                    s['checks'][f'{col(c+2)}{rr}']=sum(r['area']*r['rate'] for r in subset)
                    setcell(s,f'{col(c+1)}{rr}',f'=IF({col(c)}{rr}=0,"",{col(c+2)}{rr}/{col(c)}{rr})')
        s['last']=7+size
        s['formats']=[(f'F8:{col(nc)}{s["last"]}','0.0000;-0.0000;;@')]+[(f'{col(9+k*3)}8:{col(9+k*3)}{s["last"]}','0.00;-0.00;;@') for k in range(len(cats))]
        s['table_last']=s['last']
        for note in [fee['basis'],fee['context'],fee['note']]:
            s['last']+=1;merge(s,f'A{s["last"]}:{col(nc)}{s["last"]}',note);s['heights'][str(s['last'])]=32
    sheets.extend([s,detail])
    for s in sheets:
        s.setdefault('table_last',s['last'])
        for v in s['cells'].values():
            if isinstance(v,str) and v.startswith('=') and len(v)>8000: raise ValueError('Formula exceeds safe Excel length')
    return sheets


def style_signature(sheet, index):
    st=sheet['styles'];xf=deepcopy(st.find('s:cellXfs',NS)[index])
    for key, group in [('fontId','fonts'),('fillId','fills'),('borderId','borders')]:
        idx=int(xf.attrib.pop(key,'0'));xf.set(key,ET.tostring(st.find('s:'+group,NS)[idx],encoding='unicode'))
    nf=int(xf.attrib.get('numFmtId','0'))
    if nf>=164:
        formats=st.find('s:numFmts',NS)
        xf.set('numFmtId',next(x.attrib['formatCode'] for x in formats if int(x.attrib['numFmtId'])==nf))
    xf.attrib.pop('xfId',None)
    return ET.tostring(xf,encoding='unicode')


def verify_preserved(before, after, retained):
    for sn in retained:
        x,y=before[sn],after[sn]
        keys=set(x['cells'])|set(y['cells'])
        cache={}
        for key in keys:
            a=x['cells'].get(key,{'value':None,'formula':None,'style':0})
            b=y['cells'].get(key,{'value':None,'formula':None,'style':0})
            if a['value']!=b['value'] or a['formula']!=b['formula']: raise ValueError(f'Preserved cell changed: {sn}!{key}')
            pair=(a['style'],b['style'])
            if pair not in cache:cache[pair]=style_signature(x,pair[0])==style_signature(y,pair[1])
            if not cache[pair]:raise ValueError(f'Preserved style changed: {sn}!{key}')
        for tag in ['mergeCells','cols','pageMargins','pageSetup','printOptions']:
            aa,bb=deepcopy(x['xml'].find('s:'+tag,NS)),deepcopy(y['xml'].find('s:'+tag,NS))
            if tag=='cols':
                for node,sh in [(aa,x),(bb,y)]:
                    if node is not None:
                        for c in node:
                            if 'style' in c.attrib:c.set('style',style_signature(sh,int(c.attrib['style'])))
            if (ET.tostring(aa) if aa is not None else None)!=(ET.tostring(bb) if bb is not None else None):raise ValueError(f'Preserved sheet setting changed: {sn}/{tag}')
        def rows(sh):return [(r.attrib['r'],r.attrib.get('ht'),r.attrib.get('hidden')) for r in sh['xml'].find('s:sheetData',NS)]
        if rows(x)!=rows(y):raise ValueError(f'Preserved row layout changed: {sn}')


def validate_candidate(source, candidate, payload):
    a,b=read_book(source),read_book(candidate)
    retained=[n for n in a if n not in REPLACED_NAMES]
    if set(b)!=set(retained)|set(REPORT_NAMES):raise ValueError('Unexpected worksheet set')
    if [n for n in b if n in REPORT_NAMES] != REPORT_NAMES:raise ValueError('Unexpected report order')
    verify_preserved(a,b,retained)
    for s in payload:
        sh=b[s['name']]
        for addr,c in sh['cells'].items():
            if c['type']=='e' or (isinstance(c['value'],str) and c['value'].startswith('=')):raise ValueError(f'Formula error/text: {s["name"]}!{addr}')
        for addr,v in s['cells'].items():
            actual=sh['cells'].get(addr,{})
            if isinstance(v,str) and v.startswith('='):
                normalize_formula=lambda f:(f or '').replace("'A.1'!",'A.1!').replace("'小班明细'!",'小班明细!')
                if normalize_formula(actual.get('formula'))!=normalize_formula(v[1:]):raise ValueError(f'Formula was not preserved: {addr}')
            elif actual.get('value') not in (v, '' if v is None else v):raise ValueError(f'Cell mismatch: {addr}')
        for addr,v in s['checks'].items():
            got=val(sh,addr)
            if not isinstance(got,(int,float)) or not math.isclose(got,v,rel_tol=1e-10,abs_tol=1e-9):raise ValueError(f'Total mismatch: {s["name"]}!{addr}: {got} != {v}')


def publish(manifest):
    m=json.loads(manifest.read_text(encoding='utf-8'));source=Path(m['source']);dest=Path(m['output']);candidate=manifest.parent/'candidate.xlsx'
    if digest(source)!=m['source_hash'] or digest(candidate)!=m['candidate_hash']:raise ValueError('Source or candidate changed since validation')
    if source==dest:
        # Windows refuses replacement while the user's Excel holds a lock.
        staging=None
        try:
            with tempfile.NamedTemporaryFile(prefix='.forestry-',suffix='.xlsx',dir=dest.parent,delete=False) as f:
                staging=Path(f.name);f.write(candidate.read_bytes())
            if digest(source)!=m['source_hash']:raise ValueError('Source changed immediately before publication')
            os.replace(staging,dest)
        finally:
            if staging is not None and staging.exists():staging.unlink()
    else:
        with dest.open('xb') as f:f.write(candidate.read_bytes())
    if digest(dest)!=m['candidate_hash']:raise ValueError('Published file differs')
    print(json.dumps({'published':str(dest),'records':m['records'],'area_ha':m['area_ha'],'volume_m3':m['volume_m3']},ensure_ascii=False))


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('input',type=Path,nargs='?');ap.add_argument('--config',type=Path)
    ap.add_argument('--output',type=Path);ap.add_argument('--replace-sheets',action='store_true')
    ap.add_argument('--stage-dir',type=Path);ap.add_argument('--publish-stage',type=Path)
    args=ap.parse_args()
    if args.publish_stage:publish(args.publish_stage.resolve(strict=True));return
    if os.name!='nt':ap.error('Windows and desktop Excel are required')
    if not args.input or not args.config:ap.error('input and --config required')
    source=args.input.resolve(strict=True);output=(args.output or source).resolve()
    if source.suffix.lower()!='.xlsx' or output.suffix.lower()!='.xlsx':ap.error('Only .xlsx workbooks are supported')
    if output!=source and output.exists():ap.error('Separate output already exists')
    original=digest(source);book=read_book(source)
    if 'A.1' not in book:ap.error('A.1 worksheet missing')
    if REPLACED_NAMES.intersection(book) and not args.replace_sheets:ap.error('Existing legacy or B-series report sheets require --replace-sheets after user authorization')
    cfg=json.loads(args.config.read_text(encoding='utf-8-sig'))
    records=prepare(book,cfg);payload=build(records,cfg)
    stage=args.stage_dir.resolve() if args.stage_dir else Path(tempfile.mkdtemp(prefix='forestry-summary-'))
    if args.stage_dir:stage.mkdir(exist_ok=False,parents=True)
    candidate=stage/'candidate.xlsx'
    job={'source':str(source),'template':str(ROOT/'assets/report-headers.xlsx'),'candidate':str(candidate),'stage':str(stage),'sheets':payload,'replace_names':sorted(REPLACED_NAMES)}
    job_path=stage/'job.json';job_path.write_text(json.dumps(job,ensure_ascii=False),encoding='utf-8-sig')
    subprocess.run(['powershell.exe','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(ROOT/'scripts/excel_report.ps1'),'-Job',str(job_path)],check=True,timeout=600)
    validate_candidate(source,candidate,payload)
    if digest(source)!=original:raise ValueError('Source changed during generation')
    m={'source':str(source),'output':str(output),'source_hash':original,'candidate_hash':digest(candidate),'records':len(records),'area_ha':expected(records,'area'),'volume_m3':expected(records,'volume')}
    manifest=stage/'stage.json';manifest.write_text(json.dumps(m,ensure_ascii=False,indent=2),encoding='utf-8')
    blanks=[h for idx,h in [(29,'建设内容'),(30,'工程类别')] if any(r['values'][idx] in (None,'') for r in records)]
    print(json.dumps({**m,'stage_manifest':str(manifest),'review_pdfs':[str(stage/f'{name}.pdf') for name in REPORT_NAMES],'blank_fields':blanks,'next':'Review all print PDFs, then --publish-stage'},ensure_ascii=False))


if __name__=='__main__':main()
