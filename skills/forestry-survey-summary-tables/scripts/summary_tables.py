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
REPORT_NAMES = ['现状调查表', 'A.1'] + [f'B.{i}' for i in range(1, 8)]
INPUT_NAMES = ['原始因子表', '外业调查表', '内业处理表']
REPLACED_NAMES = set(REPORT_NAMES) | {'小班明细'} | {str(i) for i in range(1, 8)}
NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
TYPES = ['防护林林地', '特用林林地', '用材林林地', '经济林林地', '能源林林地', '苗圃地', '其他林地']
CLASSES = ['乔木林地', '竹林地', '国家特别规定灌木林地', '一般灌木林地', '疏林地', '未成林造林地', '苗圃地', '采伐迹地', '火烧迹地', '其他无立木林地']
FORESTS = ['国家级一级公益林', '国家级二级公益林', '国家级三级公益林', '省级公益林', '市县级公益林', '重点商品林', '一般商品林']
ECO = ['自然保护区', '森林公园', '湿地公园', '风景名胜区']


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
                cells[c.attrib['r']] = {'value': value, 'formula': f.text if f is not None else None, 'shared_formula': f is not None and f.attrib.get('t') == 'shared', 'type': t, 'style': int(c.attrib.get('s', 0))}
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


# Canonical internal-input schema; report formulas resolve actual column positions.
INPUT_HEADERS = ['市','县','乡','村','林班号(村小组)','小班号林地落界','小班编号','面积hm2',
 '原始地类','现状地类','林地类型','林地权属','森林类别','公益林事权','林地保护等级','林种',
 '起源','重点生态区域','原优势树种','优势树种','龄组(产期)','林龄','平均树高m','平均胸径cm',
 '郁闭度(覆盖度)','每公顷蓄积m3','小班蓄积m3','经济(竹)林株数','坡向','坡位','坡度',
 '土壤类型','土层厚度','建设内容','工程类别','使用林地性质','备注']
HEBEI_BASIS = '财税〔2015〕122号、冀财税〔2016〕25号'
HEBEI_NOTE = ('河北省参考标准：郁闭度0.2以上乔木林地（含采伐迹地、火烧迹地）、竹林地、苗圃地10元/平方米；'
 '灌木林地、疏林地、未成林造林地6元/平方米。国家级、省级公益林加倍；城市规划区内再加倍；'
 '规划区外经营性项目加倍。宜林地3元/平方米暂不纳入本表支持地类。')


def input_columns(sheet):
    heads = {normal(c['value']): a.rstrip('0123456789') for a,c in sheet['cells'].items()
             if a.endswith('3') and a.rstrip('0123456789')+'3' == a and c['value']}
    missing = [h for h in INPUT_HEADERS if normal(h) not in heads]
    if missing: raise ValueError(f'Unexpected 内业处理表 headers: {missing}')
    # Header-driven mapping also accepts the current project's optional 旧编号.
    return [heads[normal(h)] for h in INPUT_HEADERS]


def forest_display(category, authority):
    if category in ('重点公益林','一般公益林'):
        return category[:2] + {'国家公益林':'国家','地方公益林':'地方'}.get(authority,'?') + '公益林'
    return category


def prepare(book, cfg):
    if '内业处理表' not in book:
        raise ValueError('内业处理表 is required; legacy 33-column A.1 cannot be used as the new input')
    if not text(cfg.get('project_name')) or cfg.get('use_nature') not in ('临时','永久'):
        raise ValueError('Confirm project_name and use_nature')
    cfg['forest_map'] = {'重点商品林':'重点商品林','一般商品林':'一般商品林',**cfg.get('forest_map',{})}
    if any(v not in FORESTS for v in cfg['forest_map'].values()):raise ValueError('Invalid forest_map target')
    cfg['_columns'] = input_columns(book['内业处理表'])
    rows = sorted({int(''.join(filter(str.isdigit,a))) for a,c in book['内业处理表']['cells'].items()
                   if c['value'] not in (None,'') and int(''.join(filter(str.isdigit,a)))>=4})
    records=[];seen=set()
    for rr in rows:
        v=[val(book['内业处理表'],f'{c}{rr}') for c in cfg['_columns']]
        if all(x in (None,'') for x in v):continue
        ident=text(v[6])
        if not ident or ident in seen:raise ValueError(f'Missing/duplicate parcel ID: {ident}')
        seen.add(ident)
        lc,lt=text(v[9]),text(v[10]);fc=forest_display(text(v[12]),text(v[13]))
        level={'1':'Ⅰ','2':'Ⅱ','3':'Ⅲ','4':'Ⅳ','I':'Ⅰ','II':'Ⅱ','III':'Ⅲ','IV':'Ⅳ'}.get(text(v[14]),text(v[14]))
        if lc not in CLASSES or lt not in TYPES+['薪炭林林地'] or fc not in cfg['forest_map'] or level not in ['Ⅰ','Ⅱ','Ⅲ','Ⅳ']:
            raise ValueError(f'Unresolved classification: parcel {ident}: {lc}/{lt}/{fc}/{level}')
        if any(not text(x) for x in v[:4]) or v[11] not in ('国有','集体'):
            raise ValueError(f'Incomplete administration or unsupported ownership for the status-form template: {ident}')
        if text(v[35]) not in ('',cfg['use_nature']):raise ValueError(f'Use nature conflict: {ident}')
        area=number(v[7],f'occupied area {ident}',True)
        vol=number(v[26],f'approved tree stock {ident}') if lc=='乔木林地' else 0
        records.append(dict(id=ident,source=rr,values=v,admin=tuple(map(text,v[:4])),own=v[11],land_class=lc,
                            land_type=lt,forest_category=fc,level=level,area=area,volume=vol))
    if not records:raise ValueError('No input records')
    eco=cfg.get('ecology',{})
    if eco.get('mode')=='none':
        if eco.get('confirmed') is not True or any(text(r['values'][17]) not in ('','其他区域','不涉及','无') for r in records):
            raise ValueError('Confirm ecological non-involvement; A.1 must not hide an involved region')
    elif eco.get('mode')=='records':
        covered=set();keys=set()
        for row in eco.get('records',[]):
            ids=row.get('parcel_ids',[]);key=(row.get('kind'),row.get('name'))
            if row.get('kind') not in ECO or not text(row.get('name')) or not text(row.get('level')) or key in keys:
                raise ValueError('Invalid ecological record')
            if not ids or len(set(ids))!=len(ids) or set(ids)-seen:raise ValueError('Invalid ecological parcel IDs')
            covered.update(ids);keys.add(key)
        if not keys:raise ValueError('Empty ecological records')
        if any(r['id'] not in covered and text(r['values'][17]) not in ('','其他区域','不涉及','无') for r in records):
            raise ValueError('Unmapped ecological involvement')
    else:raise ValueError('ecology mode must be confirmed')
    fee=cfg.get('fee',{})
    if fee.get('mode')=='not_applicable':
        if not text(fee.get('basis')):raise ValueError('Fee non-applicability basis required')
    else:
        if fee.get('mode') not in ('reference','charge'):raise ValueError('Fee mode required')
        if cfg['use_nature']=='临时' and fee['mode']=='charge':raise ValueError('Temporary use is reference only')
        if fee.get('region')!='河北省':raise ValueError('Fee template region mismatch: provide the new regional standard before use')
        if not isinstance(fee.get('urban_planning_area'),bool) or fee.get('project_kind') not in ('公共基础设施','公共事业','国防','经营性'):
            raise ValueError('Confirm urban_planning_area and project_kind')
        for r in records:
            if r['land_class']=='乔木林地':number(r['values'][24],f'canopy {r["id"]}')
    cfg.setdefault('status_fields',{})
    allowed={'M3','D4','M4','R4','A14','A15','A20','A22','A24','A26','A28','A30','A34'}
    if set(cfg['status_fields'])-allowed:raise ValueError('Unsupported status_fields address')
    if '现状调查表' in book:
        for address in allowed:
            cell=book['现状调查表']['cells'].get(address,{})
            if cell.get('value') is not None and not cell.get('formula') and not cell.get('shared_formula'):
                cfg['status_fields'].setdefault(address,cell['value'])
    return records


def groups(records):
    result=[()]
    def walk(prefix):
        if len(prefix)==4:return
        for child in dict.fromkeys(r['admin'][:len(prefix)+1] for r in records if r['admin'][:len(prefix)]==prefix):
            result.append(child);walk(child)
    walk(())
    return result


def make_sheet(name, head, cols):
    return dict(name=name,head=head,cols=cols,cells={},merges=[],formats=[],heights={},checks={})


def merge(s, address, value):
    s['merges'].append(address);s['cells'][address.split(':')[0]]=value


def expected(records, metric):
    return float(sum((Decimal(str(r[metric])) for r in records),Decimal(0)))


def literal(value):
    return '"'+str(value).replace('"','""')+'"'


def link(column,row):
    ref=f"'内业处理表'!{column}{row}"
    return f'IF({ref}="","",{ref})'


class Totals:
    """SUMIFS selects live classifications, and excludes every subtotal via parcel ID."""
    def __init__(self,records,last):self.records=records;self.last=last
    def formula(self,metric='area',criteria=()):
        metric_col='H' if metric=='area' else 'W'
        admin=[(ord(c)-65,v) for c,v in criteria if c in 'ABCD']
        members=[r for r in self.records if all(r['admin'][i]==v for i,v in admin)]
        if not members:return '0'
        first=min(r['detail'] for r in members);last=max(r['detail'] for r in members)
        rng=lambda c:f"'A.1'!${c}${first}:${c}${last}"
        args=[rng(metric_col),rng('G'),'"<>"']
        for c,v in criteria:
            if c not in 'ABCD':args.extend([rng(c),literal(v)])
        return 'SUMIFS('+','.join(args)+')'
    def put(self,s,addr,metric='area',criteria=()):
        s['cells'][addr]='='+self.formula(metric,criteria)
        keys={'A':0,'B':1,'C':2,'D':3}
        def match(r):
            for c,v in criteria:
                actual=r['admin'][keys[c]] if c in keys else r[{'K':'own','I':'land_class','J':'land_type','L':'forest_category','M':'level','G':'id'}[c]]
                if actual!=v:return False
            return True
        s['checks'][addr]=expected([r for r in self.records if match(r)],metric)
    def union(self,s,addr,metric,base,column,values):
        s['cells'][addr]='='+('+'.join(self.formula(metric,base+[(column,v)]) for v in values) or '0')
        # Numeric expectations are calculated independently, not from generated formulas.
        checks=[]
        for v in values:
            dummy=make_sheet('',0,0);self.put(dummy,'A1',metric,base+[(column,v)]);checks.append(dummy['checks']['A1'])
        s['checks'][addr]=sum(checks)


def labels(s,gs,owners,start,shift):
    size=len(owners)
    for i,g in enumerate(gs):
        rr=start+i*size;label=['总计','合计','小计','计',g[-1] if g else ''][len(g)]
        merge(s,f'{col(shift+min(len(g),3))}{rr}:{col(shift+3)}{rr+size-1}',label)
        for j,o in enumerate(owners):s['cells'][f'{col(shift+4)}{rr+j}']=o
        if g and len(g)<4:
            end=i
            while end+1<len(gs) and gs[end+1][:len(g)]==g:end+=1
            merge(s,f'{col(shift+len(g)-1)}{rr}:{col(shift+len(g)-1)}{start+(end+1)*size-1}',g[-1])


def base_criteria(g,owner='计'):
    return list(zip('ABCD',g))+([('K',owner)] if owner!='计' else [])


def report_title(cfg, label, table=None):
    title = f'{text(cfg["project_name"])} {cfg["use_nature"]}使用林地 {label}'
    return title + (f'（表{table}）' if table else '')


def build(records,cfg):
    columns=cfg['_columns'];gs=groups(records)
    detail=make_sheet('A.1',3,28);detail['last']=3+len(gs)+len(records)
    detail['cells']['A1']=report_title(cfg,'小班因子调查表','A.1')
    totals=Totals(records,detail['last']);rr=4
    # Administrative layout is rebuilt on regrouping; classifications stay live.
    group_rows={}
    for g in gs:
        items=[r for r in records if r['admin'][:len(g)]==g]
        merge(detail,f'{col(min(len(g)+1,5))}{rr}:G{rr}',['省总计','市合计','县合计','镇合计','小计'][len(g)])
        group_rows[g]=rr
        if g:detail['cells'][f'{col(len(g))}{rr}']=g[-1]
        detail['heights'][str(rr)]=18
        # Administrative leaf ranges assigned after all detail rows are known.
        detail.setdefault('subtotals',[]).append((rr,items));rr+=1
        if len(g)!=4:continue
        for r in items:
            r['detail']=rr;sr=r['source']
            indices=[4,5,6,7,9,10,11,12,14,15,16,17,19,20,22,23,24,25,26,27,33,34,35,36]
            for dc,idx in enumerate(indices,5):
                expr=link(columns[idx],sr)
                if dc==8:
                    ref=f"'内业处理表'!{columns[7]}{sr}"
                    expr=f'IF(AND(ISNUMBER({ref}),{ref}>0),{ref},NA())'
                if dc==12:
                    cat=f"'内业处理表'!{columns[12]}{sr}";authority=f"'内业处理表'!{columns[13]}{sr}"
                    expr=f'IF({cat}="","",IF(OR({cat}="重点公益林",{cat}="一般公益林"),IF(OR({authority}="国家公益林",{authority}="地方公益林"),LEFT({cat},2)&LEFT({authority},2)&"公益林","待核定公益林事权"),{cat}))'
                if dc==13:
                    ref=f"'内业处理表'!{columns[14]}{sr}"
                    expr=link(columns[14],sr)
                    for key,value in [('1','Ⅰ'),('2','Ⅱ'),('3','Ⅲ'),('4','Ⅳ'),('I','Ⅰ'),('II','Ⅱ'),('III','Ⅲ'),('IV','Ⅳ')]:
                        expr=f'IF({ref}&""={literal(key)},{literal(value)},{expr})'
                if dc==16:
                    ref=f"'内业处理表'!{columns[17]}{sr}"
                    expr=f'IF(OR({ref}="",{ref}="其他区域",{ref}="无"),"不涉及",{ref})'
                if dc==23:
                    ref=f"'内业处理表'!{columns[26]}{sr}"
                    expr=f'IF(I{rr}<>"乔木林地","",IF(AND(ISNUMBER({ref}),{ref}>=0),{ref},NA()))'
                detail['cells'][f'{col(dc)}{rr}']='='+expr
            rr+=1
    for g,row in group_rows.items():
        if g:
            end=max(r['detail'] for r in records if r['admin'][:len(g)]==g)
            detail['merges'].append(f'{col(len(g))}{row}:{col(len(g))}{end}')
    for row,items in detail.pop('subtotals'):
        for c,metric in [('H','area'),('W','volume')]:
            first=min(r['detail'] for r in items);last=max(r['detail'] for r in items)
            detail['cells'][f'{c}{row}']=f'=SUMIFS({c}{first}:{c}{last},G{first}:G{last},"<>")'
            detail['checks'][f'{c}{row}']=expected(items,metric)
    detail['formats']=[(f'H4:H{rr-1}','0.0000;-0.0000;;@'),(f'S4:U{rr-1}','0.0;-0.0;;@'),(f'W4:W{rr-1}','0.0;-0.0;;@')]
    owners=['计','国有','集体'];size=len(gs)*3;sheets=[detail]
    titles=['林地类型','现状地类','森林类别','林地保护等级']
    for n in range(1,5):
        head,nc=(7,17) if n==3 else (5,14 if n==1 else 17 if n==2 else 11)
        s=make_sheet(f'B.{n}',head,nc);s['cells']['A1']=report_title(cfg,f'按{titles[n-1]}面积与蓄积统计表',f'B.{n}')
        if n==2:s['cells']['Q5']='其他无立\n木林地'
        for metric,offset in [('area',0),('volume',size)]:
            start=head+1+offset;labels(s,gs,owners,start,2);merge(s,f'A{start}:A{start+size-1}','面积' if metric=='area' else '蓄积')
            for i,g in enumerate(gs):
                for j,o in enumerate(owners):
                    row=start+i*3+j;base=base_criteria(g,o)
                    if n==3:
                        buckets=[FORESTS,FORESTS[:5],FORESTS[:3],FORESTS[:1],FORESTS[1:2],FORESTS[2:3],FORESTS[3:4],FORESTS[4:5],FORESTS[5:],FORESTS[5:6],FORESTS[6:]]
                        for c,leaves in enumerate(buckets,7):totals.union(s,f'{col(c)}{row}',metric,base,'L',[k for k,v in cfg['forest_map'].items() if v in leaves])
                    else:
                        totals.put(s,f'G{row}',metric,base)
                        for c,v in enumerate(TYPES if n==1 else CLASSES if n==2 else ['Ⅰ','Ⅱ','Ⅲ','Ⅳ'],8):
                            key='J' if n==1 else 'I' if n==2 else 'M'
                            if n==1 and v=='能源林林地':totals.union(s,f'{col(c)}{row}',metric,base,key,[v,'薪炭林林地'])
                            else:totals.put(s,f'{col(c)}{row}',metric,base+[(key,v)])
            s['formats'].append((f'G{start}:{col(nc)}{start+size-1}','0.0000;-0.0000;;@' if metric=='area' else '0.0;-0.0;;@'))
        s['last']=head+2*size;sheets.append(s)
    classes=CLASSES[:5]+[c for c in CLASSES[5:] if any(r['land_class']==c for r in records)]
    nc=3+2*(1+len(classes));s=make_sheet('B.5',6,nc)
    merge(s,f'A1:{col(nc)}2',report_title(cfg,'分森林类别按现状地类面积统计表','B.5'))
    merge(s,f'A3:{col(nc)}3','单位：公顷');merge(s,'A4:A5','县、林业局');merge(s,'B4:C5','森林类别')
    s['cells']['A6']=1;merge(s,'B6:C6',2)
    for k,lc in enumerate(['合计']+classes):
        c=4+k*2;merge(s,f'{col(c)}4:{col(c+1)}4',lc)
        for j,o in enumerate(owners[1:]):s['cells'][f'{col(c+j)}5']=o;s['cells'][f'{col(c+j)}6']=c+j-1
    counties=list(dict.fromkeys(r['admin'][:2] for r in records))
    leaves=[FORESTS[:5],FORESTS[:3],FORESTS[3:4],FORESTS[4:5],FORESTS[5:],FORESTS[5:6],FORESTS[6:]]
    for i,g in enumerate([()]+counties):
        start=7+7*i;name='合计' if not g else ''.join(g) if sum(c[1]==g[1] for c in counties)>1 else g[1]
        merge(s,f'A{start}:A{start+6}',name);merge(s,f'B{start}:B{start+3}','公益林地');merge(s,f'B{start+4}:B{start+6}','商品林地')
        for j,label in enumerate(['小计','国家级公益林地','省级公益林地','市县级公益林地','小计','重点商品林地','一般商品林地']):
            row=start+j;s['cells'][f'C{row}']=label
            for k,lc in enumerate([None]+classes):
                for h,o in enumerate(owners[1:]):
                    base=base_criteria(g,o)+([('I',lc)] if lc else [])
                    totals.union(s,f'{col(4+k*2+h)}{row}','area',base,'L',[k for k,v in cfg['forest_map'].items() if v in leaves[j]])
    s['last']=6+7*(1+len(counties));s['formats']=[(f'A6:{col(nc)}6','0'),(f'D7:{col(nc)}{s["last"]}','0.0000;-0.0000;;@')];sheets.append(s)
    s=make_sheet('B.6',5,11);s['cells']['A1']=report_title(cfg,'重点生态区域林地面积统计表','B.6')
    if cfg['ecology']['mode']=='none':
        s['cells']['A6']='=IF(COUNTIF(\'A.1\'!P4:P'+str(detail['last'])+',"不涉及")=COUNTA(\'内业处理表\'!'+columns[6]+'4:'+columns[6]+str(max(r['source'] for r in records))+'),"不涉及","需核查生态区域")';s['last']=6
    else:
        for row,reg in enumerate(cfg['ecology']['records'],6):
            s['cells'][f'A{row}']=reg['name'];totals.union(s,f'B{row}','area',[],'G',reg['parcel_ids'])
            c=3+2*ECO.index(reg['kind']);s['cells'][f'{col(c)}{row}']=reg['level'];totals.union(s,f'{col(c+1)}{row}','area',[],'G',reg['parcel_ids'])
        s['last']=5+len(cfg['ecology']['records'])
    s['formats']=[(f'B6:J{s["last"]}','0.0000;-0.0000;;@')];sheets.append(s)
    sheets.append(build_fee(records,cfg,gs,totals))
    status=make_sheet('现状调查表',35,19);status['last']=35;status['preserve_body']=True
    status['cells'].update(cfg['status_fields']);status['cells'].update({'A1':report_title(cfg,'现状调查表'),'D3':cfg['project_name'],'A16':report_title(cfg,'现状调查表（续）')})
    for metric,start in [('area',7),('volume',10)]:
        for j,o in enumerate(owners):
            row=start+j;base=base_criteria((),o);totals.put(status,f'D{row}',metric,base)
            for c,lt in enumerate(TYPES[:6],5):
                totals.union(status,f'{col(c)}{row}',metric,base,'J',[lt,'薪炭林林地'] if lt=='能源林林地' else [lt])
            for c,lc in enumerate(CLASSES[:9],11):totals.put(status,f'{col(c)}{row}',metric,base+[('I',lc)])
        status['formats'].append((f'D{start}:S{start+2}','0.0000;-0.0000;;@' if metric=='area' else '0.0;-0.0;;@'))
    # The fixed status-form has no column for 其他林地. Do not silently hide it.
    if any(r['land_class']==CLASSES[-1] or r['land_type']=='其他林地' for r in records):
        raise ValueError('The status-form template has no other-land column; confirm an extended form before generating')
    for row in [7,8,9,10,11,12]:
        # Surface later edits to unsupported categories without declaring success.
        metric='area' if row<10 else 'volume';base=base_criteria((),owners[(row-7)%3])
        check=totals.formula('area',base)
        covered='+'.join(totals.formula('area',base+[('I',c)]) for c in CLASSES[:9])
        old=status['cells'][f'D{row}'][1:];status['cells'][f'D{row}']=f'=IF(ABS({check}-({covered}))<0.00000001,{old},"分类需核查")'
    if cfg['fee']['mode']=='not_applicable':status['cells']['A13']='="森林植被恢复费："&\'B.7\'!A8'
    else:status['cells']['A13']='=IF(ISNUMBER(\'B.7\'!G8),"森林植被恢复费'+('参考测算' if cfg['fee']['mode']=='reference' else '测算')+'："&TEXT(\'B.7\'!G8,"0.0000")&"万元","森林植被恢复费：待核定")'
    sheets.insert(0,status)
    for s in sheets:
        s.setdefault('table_last',s['last'])
        for v in s['cells'].values():
            if isinstance(v,str) and v.startswith('=') and len(v)>8000:raise ValueError('Formula too long; reduce workbook scope')
    return sheets


def build_fee(records,cfg,gs,totals):
    fee=cfg['fee'];s=make_sheet('B.7',7,13)
    s['cells']['A3']='单位：面积—公顷；基础单价—元/平方米；金额—万元'
    s['cells']['A1']=report_title(cfg,'森林植被恢复费'+('参考测算表' if fee['mode']=='reference' else '测算表'),'B.7')
    if fee['mode']=='not_applicable':
        merge(s,'A8:M9','不适用：'+fee['basis']);s['last']=9;return s
    s['cells'].update({'H4':'乔木林地等（10元类）','K4':'灌木林地等（6元类）','I8':10,'L8':6})
    # Global project conditions are explicit; forest status remains a live predicate.
    factor=2 if fee['urban_planning_area'] or fee['project_kind']=='经营性' else 1
    def area_expr(base,band,forests=None):
        tests=[[('I','乔木林地'),('U','>=0.2')]]+[[('I',c)] for c in ['竹林地','苗圃地','采伐迹地','火烧迹地']] if band==10 else [[('I',c)] for c in ['国家特别规定灌木林地','一般灌木林地','疏林地','未成林造林地']]
        return '+'.join(totals.formula('area',base+t+([('L',f)] if f else [])) for t in tests for f in (forests if forests is not None else [None])) or '0'
    labels(s,gs,['计','国有','集体'],8,1)
    doubled=[k for k,v in cfg['forest_map'].items() if v in FORESTS[:4]]
    for i,g in enumerate(gs):
        for j,o in enumerate(['计','国有','集体']):
            row=8+3*i+j;base=base_criteria(g,o);totals.put(s,f'F{row}','area',base)
            s['cells'][f'H{row}']='='+area_expr(base,10);s['cells'][f'K{row}']='='+area_expr(base,6)
            known='+'.join(totals.formula('area',base+[('L',k)]) for k in cfg['forest_map'])
            guard=f'AND(ABS(F{row}-H{row}-K{row})<0.00000001,ABS(F{row}-({known}))<0.00000001,ISNUMBER($I$8),ISNUMBER($L$8),$I$8>=0,$L$8>=0)'
            for band,ac,rate,fc in [(10,'H','I','J'),(6,'K','L','M')]:
                if row!=8:s['cells'][f'{rate}{row}']=f'=${rate}$8'
                extra=area_expr(base,band,doubled)
                s['cells'][f'{fc}{row}']=f'=IF({guard},({ac}{row}+({extra}))*${rate}$8*{factor},"待核定")'
            s['cells'][f'G{row}']=f'=IF(AND(ISNUMBER(J{row}),ISNUMBER(M{row})),J{row}+M{row},"待核定")'
            items=[r for r in records if r['admin'][:len(g)]==g and (o=='计' or r['own']==o)]
            amount=0;supported=True
            for r in items:
                lc=r['land_class'];rate=10 if lc in ['竹林地','苗圃地','采伐迹地','火烧迹地'] or lc=='乔木林地' and r['values'][24]>=.2 else 6 if lc in ['国家特别规定灌木林地','一般灌木林地','疏林地','未成林造林地'] else None
                if rate is None:supported=False;continue
                amount+=r['area']*rate*factor*(2 if r['forest_category'] in doubled else 1)
            if supported:s['checks'][f'G{row}']=amount
    s['last']=7+len(gs)*3;s['table_last']=s['last'];s['last']+=1
    note=HEBEI_BASIS+'。\n'+HEBEI_NOTE+'\n'+('规划区内' if fee['urban_planning_area'] else '规划区外')+'；'+fee['project_kind']+'。'
    if fee['mode']=='reference':note+='本表仅为参考测算，不作为应缴金额；临时使用林地不再收取森林植被恢复费，实际恢复工程费用另计（国家林草局2023年第13号公告）。'
    merge(s,f'A{s["last"]}:M{s["last"]}',note);s['heights'][str(s['last'])]=100
    s['formats']=[(f'F8:M{s["table_last"]}','0.0000;-0.0000;;@'),(f'I8:I{s["table_last"]}','0.00;-0.00;;@'),(f'L8:L{s["table_last"]}','0.00;-0.00;;@')]
    return s


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


def template_fonts(payload, template):
    """Restore explicit fonts after cross-workbook copying of named Excel styles."""
    for s in payload:
        ref=template[s['name']];styles=ref['styles']
        def point(address):
            letters=address.rstrip('0123456789');n=0
            for char in letters:n=n*26+ord(char)-64
            return n,int(address[len(letters):])
        merged=[]
        for node in ref['xml'].findall('s:mergeCells/s:mergeCell',NS):
            a,b=node.attrib['ref'].split(':');merged.append((*point(a),*point(b),a))
        groups={}
        for row in range(1,s['last']+1):
            rr=row if row<=s['head'] else 9 if s['name']=='A.1' else s['head']+1
            run=None;start=1
            for c in range(1,s['cols']+2):
                signature=None
                if c<=s['cols']:
                    cc=4 if s['name']=='B.5' and c>3 else c
                    address=next((a for x1,y1,x2,y2,a in merged if x1<=cc<=x2 and y1<=rr<=y2),f'{col(cc)}{rr}')
                    idx=ref['cells'].get(address,{}).get('style',0)
                    xf=styles.find('s:cellXfs',NS)[idx]
                    font=styles.find('s:fonts',NS)[int(xf.attrib.get('fontId',0))]
                    def attr(tag,default):
                        node=font.find('s:'+tag,NS)
                        return node.attrib.get('val',default) if node is not None else default
                    signature=(attr('name','宋体'),float(attr('sz',8)),font.find('s:b',NS) is not None,font.find('s:i',NS) is not None)
                if c==1:run=signature
                elif signature!=run:
                    groups.setdefault(run,[]).append(f'{col(start)}{row}:{col(c-1)}{row}')
                    start=c;run=signature
        s['font_formats']=[]
        for signature,ranges in groups.items():
            chunk=[]
            for address in ranges:
                if len(','.join(chunk+[address]))>200:
                    s['font_formats'].append([','.join(chunk),*signature]);chunk=[]
                chunk.append(address)
            if chunk:s['font_formats'].append([','.join(chunk),*signature])


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
            if tag=='mergeCells':
                for node in (aa,bb):
                    if node is not None:node[:]=sorted(node,key=lambda n:n.attrib['ref'])
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
    expected_order=[n for n in INPUT_NAMES if n in a]+REPORT_NAMES
    if list(b)[:len(expected_order)] != expected_order:raise ValueError('Unexpected input/report order')
    verify_preserved(a,b,retained)
    for s in payload:
        sh=b[s['name']]
        for addr,c in sh['cells'].items():
            if c['type']=='e' or (isinstance(c['value'],str) and c['value'].startswith('=')):raise ValueError(f'Formula error/text: {s["name"]}!{addr}')
        for addr,v in s['cells'].items():
            actual=sh['cells'].get(addr,{})
            if isinstance(v,str) and v.startswith('='):
                normalize_formula=lambda f:(f or '').replace("'A.1'!",'A.1!').replace("'内业处理表'!",'内业处理表!').replace("'B.7'!",'B.7!')
                # Native Excel already verifies every expanded formula on reopen.
                # XML shared-formula followers intentionally omit formula text.
                if not actual.get('shared_formula') and normalize_formula(actual.get('formula'))!=normalize_formula(v[1:]):raise ValueError(f'Formula was not preserved: {addr}')
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
    if '内业处理表' not in book:ap.error('内业处理表 worksheet missing; finalize it manually before report generation')
    if REPLACED_NAMES.intersection(book) and not args.replace_sheets:ap.error('Existing legacy or B-series report sheets require --replace-sheets after user authorization')
    cfg=json.loads(args.config.read_text(encoding='utf-8-sig'))
    records=prepare(book,cfg);payload=build(records,cfg)
    template_fonts(payload,read_book(ROOT/'assets/report-headers.xlsx'))
    stage=args.stage_dir.resolve() if args.stage_dir else Path(tempfile.mkdtemp(prefix='forestry-summary-'))
    if args.stage_dir:stage.mkdir(exist_ok=False,parents=True)
    candidate=stage/'candidate.xlsx'
    job={'source':str(source),'template':str(ROOT/'assets/report-headers.xlsx'),'candidate':str(candidate),'stage':str(stage),'sheets':payload,'replace_names':sorted(REPLACED_NAMES),'input_names':[n for n in INPUT_NAMES if n in book]}
    job_path=stage/'job.json';job_path.write_text(json.dumps(job,ensure_ascii=False),encoding='utf-8-sig')
    subprocess.run(['powershell.exe','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(ROOT/'scripts/excel_report.ps1'),'-Job',str(job_path)],check=True,timeout=600)
    validate_candidate(source,candidate,payload)
    if digest(source)!=original:raise ValueError('Source changed during generation')
    m={'source':str(source),'output':str(output),'source_hash':original,'candidate_hash':digest(candidate),'records':len(records),'area_ha':expected(records,'area'),'volume_m3':expected(records,'volume')}
    manifest=stage/'stage.json';manifest.write_text(json.dumps(m,ensure_ascii=False,indent=2),encoding='utf-8')
    blanks=[h for idx,h in [(33,'建设内容'),(34,'工程类别')] if any(r['values'][idx] in (None,'') for r in records)]
    print(json.dumps({**m,'stage_manifest':str(manifest),'review_pdfs':[str(stage/f'{name}.pdf') for name in REPORT_NAMES],'blank_fields':blanks,'next':'Review all print PDFs, then --publish-stage'},ensure_ascii=False))


if __name__=='__main__':main()
