# 이카운트 엑셀(스타일 XML 깨짐) 읽기: styles.xml 무시하고 값만 읽음
import zipfile, re, sys
from xml.etree import ElementTree as ET
NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
def read_rows(path):
    z=zipfile.ZipFile(path)
    ss=[]
    if 'xl/sharedStrings.xml' in z.namelist():
        for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si',NS):
            ss.append(''.join(t.text or '' for t in si.iter('{%s}t'%NS['m'])))
    sheet=sorted(n for n in z.namelist() if n.startswith('xl/worksheets/sheet'))[0]
    rows=[]
    for row in ET.fromstring(z.read(sheet)).iter('{%s}row'%NS['m']):
        r={}
        for c in row.findall('m:c',NS):
            col=re.match(r'[A-Z]+',c.get('r')).group(); t=c.get('t'); v=c.find('m:v',NS)
            if t=='inlineStr': val=''.join(x.text or '' for x in c.iter('{%s}t'%NS['m']))
            elif v is None or v.text is None: val=None
            elif t=='s': val=ss[int(v.text)]
            elif t in ('str','b','e'): val=v.text
            else:
                f=float(v.text); val=int(f) if f.is_integer() else f
            r[col]=val
        if r:
            n=max(sum((ord(ch)-64)*26**i for i,ch in enumerate(reversed(k))) for k in r)
            out=[None]*n
            for k,val in r.items(): out[sum((ord(ch)-64)*26**i for i,ch in enumerate(reversed(k)))-1]=val
            rows.append(out)
    return rows
if __name__=='__main__':
    for p in sys.argv[1:]:
        rows=read_rows(p); print('=====',p[-40:],len(rows))
        for r in rows[:12]: print(r)
        print('...'); 
        for r in rows[-4:]: print(r)
