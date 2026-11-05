#!/usr/bin/env python3
"""Independent XLSM-to-generated-HTML numerical audit. Read-only; no statistics in output."""
import argparse, json, re
from collections import Counter
from pathlib import Path
import openpyxl

KEYS=('att','yw','ym','tr','lur','minb','mins','nc','ysa','kids')
SI_COLS={'Moscow Stake':'B','St.Petersburg Stake':'G','Saratov District':'L','Rostov District':'Q','Moscow Mission':'X','Yekaterinburg Msn':'AC','Yekaterinburg District':'AH','Novosibirsk Msn':'AO','Novosibirsk District':'AT','Armenia Mission':'BA','Yerevan District':'BF','area':'BM'}
def quarter(v): return isinstance(v,str) and bool(re.fullmatch(r'\d\d-\d',v.strip()))
def number(v): return v if isinstance(v,(int,float)) and not isinstance(v,bool) else 0
def load_data(path):
    html=Path(path).read_text(encoding='utf-8')
    match=re.search(r'const DATA\s*=\s*',html)
    if not match or html[match.end():].startswith('window.__DATA__'):
        raise ValueError('Encrypted HTML: generate a temporary unencrypted report for local audit')
    return json.JSONDecoder().raw_decode(html[match.end():])[0]
def traverse(n):
    yield n
    for c in n.get('children',[]): yield from traverse(c)
def main():
    ap=argparse.ArgumentParser();ap.add_argument('xlsm');ap.add_argument('html');args=ap.parse_args()
    wb=openpyxl.load_workbook(args.xlsm,data_only=True,read_only=False)
    data=load_data(args.html); qs=data['quarters']; errors=[]; counts=Counter(); samples=[]
    def check(kind,where,actual,expected):
        counts[kind]+=1
        if actual!=expected:
            errors.append(kind)
            if len(samples)<15:samples.append((kind,where,actual,expected))
    nodes=list(traverse(data['tree']))
    leaves=[n for n in nodes if not n.get('children')]
    sheet_rows={}
    for n in leaves:
        sheet,unit=n['id'].split('/',1)
        ws=wb[sheet]
        if sheet not in sheet_rows:
            blocks=[];r=1
            while r<=ws.max_row:
                if quarter(ws.cell(r,1).value) and (r==1 or not quarter(ws.cell(r-1,1).value)):
                    rows={}
                    while r<=ws.max_row and quarter(ws.cell(r,1).value):
                        rows[ws.cell(r,1).value.strip()]=r;r+=1
                    blocks.append(rows)
                else:r+=1
            sheet_rows[sheet]=blocks
            check('block_count',sheet,len(blocks),10)
        headers={ws.cell(1,c).value.strip():c for c in range(2,ws.max_column+1) if isinstance(ws.cell(1,c).value,str) and ws.cell(1,c).value.strip()}
        if unit not in headers:
            errors.append('missing_column');continue
        c=headers[unit]
        for ki,k in enumerate(KEYS):
            for qi,q in enumerate(qs):
                row=sheet_rows[sheet][ki].get(q)
                expected=[number(ws.cell(row,c).value),number(ws.cell(row,c+1).value)] if row else [0,0]
                check('leaf_pair',f'{sheet}/{unit} {k} {q}',n['m'][k][qi],expected)
    for n in nodes:
        if not n.get('children'):continue
        for k in KEYS:
            for i,q in enumerate(qs):
                expected=[sum(ch['m'][k][i][j] for ch in n['children']) for j in (0,1)]
                check('aggregate_pair',f'{n["id"]} {k} {q}',n['m'][k][i],expected)
    for n in nodes:
        if n['id'] not in SI_COLS:continue
        ws=wb['S&I Summary'];c=openpyxl.utils.column_index_from_string(SI_COLS[n['id']]);
        starts={}
        for row in range(1,ws.max_row+1):
            v=ws.cell(row,2).value
            if isinstance(v,str) and v.strip() in ('Youth Potential','YSA Potential'):
                starts.setdefault(v.strip(),row+1)
        for k,marker,off in [('sem','Youth Potential',1),('inst','YSA Potential',1),('semS','Youth Potential',2),('instS','YSA Potential',2)]:
            if marker not in starts:continue
            rows={};r=starts[marker]
            while r<=ws.max_row and quarter(ws.cell(r,1).value):
                rows[ws.cell(r,1).value.strip()]=r;r+=1
            for i,q in enumerate(qs):
                rr=rows.get(q); expected=[number(ws.cell(rr,c+off).value),number(ws.cell(rr,c).value)] if rr else [0,0]
                check('si_pair',f'{n["id"]} {k} {q}',n['m'][k][i],expected)
    print('QUARTERS:',', '.join(qs))
    print('NODES:',len(nodes),'LEAVES:',len(leaves))
    for kind,n in counts.items():print(f'{kind}: {n} checks')
    print('MISMATCHES:',len(errors))
    for kind,where,a,b in samples:print('DIFF',kind,where,'html=',a,'excel=',b)
    if errors:raise SystemExit(1)
if __name__=='__main__':main()
