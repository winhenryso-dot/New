import sys, re, openpyxl, warnings
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
warnings.filterwarnings('ignore')
fa, fb, out = sys.argv[1:4]
def load(f):
    ws = openpyxl.load_workbook(f, data_only=True).active
    name = str(ws['A1'].value).replace('회사명 :','').strip()
    rows = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        if r[0] and r[1] is not None:  # 마지막 출력일시 행 제외
            rows[str(r[0]).strip()] = r
    return name, rows
na, A = load(fa); nb, B = load(fb)
UNIT = {'R/L':'RL','개':'EA','롤':'RL'}
nu = lambda u: UNIT.get(str(u).strip().upper(), str(u).strip().upper()) if u else ''
norm = lambda s: re.sub(r'\s+','',str(s or '')).upper()
common = sorted(set(A)&set(B)); onlyA = sorted(set(A)-set(B)); onlyB = sorted(set(B)-set(A))

same, diff = [], []
for c in common:
    a, b = A[c], B[c]
    probs = []
    if norm(a[1]) != norm(b[1]): probs.append('품목명')
    if norm(a[2]) != norm(b[2]): probs.append('규격')
    if a[3] != b[3]:
        probs.append('단위(표기만 다름)' if nu(a[3])==nu(b[3]) or not a[3] or not b[3] else '단위(실제 다름)⚠')
    (diff if probs else same).append((c, a, b, probs))

# 한쪽에만 있는 코드 중 같은 품목 의심 (품목명+규격 동일 or 코드 유사)
key = lambda r: (norm(r[1]), norm(r[2]))
kb = {}
for c in onlyB: kb.setdefault(key(B[c]), []).append(c)
squash = lambda c: re.sub(r'[^0-9A-Z]','',c.upper())
sq_b = {squash(c): c for c in onlyB}
sus = []; usedA=set(); usedB=set()
for c in onlyA:
    hit = None; why=''
    if squash(c) in sq_b: hit, why = sq_b[squash(c)], '코드 오타 의심(대소문자/하이픈)'
    elif key(A[c]) in kb: hit, why = kb[key(A[c])][0], '품목명+규격 동일, 코드 다름'
    if hit: sus.append((c, hit, why)); usedA.add(c); usedB.add(hit)

wb = openpyxl.Workbook()
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='305496')
def sheet(title, header, data, widths):
    ws = wb.create_sheet(title)
    ws.append(header)
    for cell in ws[1]: cell.font, cell.fill = H, HF
    for d in data: ws.append(list(d))
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions
    return ws
ws = wb.active; ws.title = '요약'
summ = [('항목','건수','설명'),
 (f'{na} 품목 수', len(A), ''), (f'{nb} 품목 수', len(B), ''),
 ('공통 코드', len(common), '두 법인에 같은 코드가 있음'),
 ('  └ 완전 일치', len(same), '바로 합산 가능'),
 ('  └ 코드는 같은데 내용 다름', len(diff), "'코드같음_내용다름' 시트 확인 → 한쪽 기준으로 통일 필요"),
 (f'{na}에만 있음', len(onlyA), ''), (f'{nb}에만 있음', len(onlyB), ''),
 ('같은 품목인데 코드가 다른 것(의심)', len(sus), "'코드다름_같은품목의심' 시트 → 코드 통일 필요")]
for r in summ: ws.append(r)
for cell in ws[1]: cell.font, cell.fill = H, HF
ws.column_dimensions['A'].width=34; ws.column_dimensions['B'].width=10; ws.column_dimensions['C'].width=60

sheet('코드같음_내용다름', ['품목코드','차이항목',f'{na} 품목명',f'{nb} 품목명',f'{na} 규격',f'{nb} 규격',f'{na} 단위',f'{nb} 단위'],
      [(c, ', '.join(p), a[1], b[1], a[2], b[2], a[3], b[3]) for c,a,b,p in diff], [14,22,28,28,34,34,8,8])
sheet('코드다름_같은품목의심', ['사유',f'{na} 코드',f'{nb} 코드','품목명','규격'],
      [(w, ca, cb, A[ca][1], A[ca][2]) for ca,cb,w in sus], [30,14,14,30,36])
cols = lambda r: (r[0], r[1], r[2], r[3], r[6], r[8])
sheet(f'{na}에만', ['품목코드','품목명','규격','단위','품목구분','그룹'], [cols(A[c]) for c in onlyA if c not in usedA], [14,30,36,8,12,14])
sheet(f'{nb}에만', ['품목코드','품목명','규격','단위','품목구분','그룹'], [cols(B[c]) for c in onlyB if c not in usedB], [14,30,36,8,12,14])
sheet('통합품목마스터(일치)', ['품목코드','품목명','규격','단위','품목구분','그룹'], [cols(a) for c,a,b,p in same], [14,30,36,8,12,14])
wb.save(out)
for r in summ[1:]: print(r[0], r[1])
import collections; print(collections.Counter(p for *_,ps in diff for p in ps))
