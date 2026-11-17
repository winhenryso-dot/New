"""법인 ERP 판매현황 엑셀 → 재고 ERP '판매입력 웹자료올리기' 엑셀 변환.

사용법:
  python3 sales_to_stock.py --namkyung 남경판매현황.xlsx --eng 이엔지판매현황.xlsx \
      [--map 품목코드_매핑.xlsx] [--stock-items 재고ERP_품목등록.xlsx] \
      [--from 2026-11-18] [--to 2026-11-18] -o 재고ERP_업로드.xlsx

규칙
 - 일자: 회계 ERP 판매일자 그대로 (이월 건도 회계 날짜로 이동, 비고는 적요에 그대로 남김)
 - 거래처명: 회계 ERP 거래처명 그대로 / 출하창고: 2차공장 / 단가·금액: 0
 - 법인 간 내부거래(상대 법인 거래처)는 제외 → '제외_내부거래' 시트
 - 단가 정정/차액 건은 업로드에서 빼고 '확인필요' 시트로
 - Z- 코드(매출할인 등 비재고 품목)는 제외
 - 반품(마이너스 수량), 무상(0원) 출고는 그대로 반영
 - 생산품만 반영: 회계 ERP 품목구분 [제품]
 - 품목코드는 매핑표가 있으면 재고 ERP 코드로 변환
"""
import argparse, re, sys, os
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ecount_xlsx import read_rows

WAREHOUSE = '2차공장'
INTERNAL = {'남경': {'남경이엔지'}, '이엔지': {'주남경', '(주)남경', '남경'}}
PRICE_FIX = re.compile(r'단가\s*(차액|수정|정정|변경)|차액분|금액\s*(수정|정정)')
TEMPLATE = ['일자', '순번', '거래처코드', '거래처명', '담당자', '출하창고', '거래유형', '통화', '환율',
            '품목코드', '품목명', '규격', '수량', '단가', '외화공급가액', '공급가액', '부가세', '적요', '생산전표생성']


def parse_date(s):
    """'26/10/01-1' 또는 '2026/10/01-1' → (date, 전표번호)"""
    m = re.match(r'(\d{2,4})/(\d{1,2})/(\d{1,2})-(\d+)', str(s).strip())
    if not m:
        return None, None
    y = int(m.group(1)); y += 2000 if y < 100 else 0
    return date(y, int(m.group(2)), int(m.group(3))), int(m.group(4))


def load_sales(path, entity):
    rows = read_rows(path)
    header = next(i for i, r in enumerate(rows) if r and r[0] == '일자')
    h = [str(x or '').strip() for x in rows[header]]
    col = lambda *names: next(h.index(n) for n in names if n in h)
    ci = dict(date=0, cust=col('검색창내용', '거래처명'), item=col('품명 및 규격'), qty=col('수량'),
              memo=col('비고', '적요'), code=col('품목코드코드', '품목코드'))
    out = []
    for r in rows[header + 1:]:
        r = r + [None] * (len(h) - len(r))
        d, no = parse_date(r[ci['date']])
        if not d:
            continue  # 월계/총합계/출력일시 행
        out.append(dict(entity=entity, date=d, slip=no, cust=(r[ci['cust']] or '').strip(),
                        item=r[ci['item']], qty=r[ci['qty']] or 0, memo=r[ci['memo']] or '',
                        code=str(r[ci['code']] or '').strip()))
    return out


def load_map(path):
    """매핑표: 법인 | 법인품목코드 | 재고ERP품목코드 (헤더 1행)"""
    m = {}
    if not path:
        return m
    ws = openpyxl.load_workbook(path, data_only=True).active
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r and r[0] and r[1] and r[2]:
            m[(str(r[0]).strip(), str(r[1]).strip())] = str(r[2]).strip()
    return m


def load_stock_items(path):
    if not path:
        return None
    rows = read_rows(path)
    header = next(i for i, r in enumerate(rows) if r and r[0] == '품목코드')
    return {str(r[0]).strip() for r in rows[header + 1:] if r and r[0]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--namkyung'); ap.add_argument('--eng')
    ap.add_argument('--map'); ap.add_argument('--stock-items')
    ap.add_argument('--items-namkyung', help='남경 회계 ERP 품목등록 (품목구분으로 생산품 판별)')
    ap.add_argument('--items-eng', help='이엔지 회계 ERP 품목등록')
    ap.add_argument('--from', dest='dfrom'); ap.add_argument('--to', dest='dto')
    ap.add_argument('-o', '--out', required=True)
    a = ap.parse_args()

    lines = []
    if a.namkyung: lines += load_sales(a.namkyung, '남경')
    if a.eng: lines += load_sales(a.eng, '이엔지')
    if a.dfrom: lines = [l for l in lines if l['date'] >= date.fromisoformat(a.dfrom)]
    if a.dto: lines = [l for l in lines if l['date'] <= date.fromisoformat(a.dto)]
    cmap, stock = load_map(a.map), load_stock_items(a.stock_items)
    kind = {}
    for e, f in (('남경', a.items_namkyung), ('이엔지', a.items_eng)):
        if f:
            rows = read_rows(f)
            hi = next(i for i, r in enumerate(rows) if r and r[0] == '품목코드')
            ki = rows[hi].index('품목구분')
            for r in rows[hi + 1:]:
                if r and r[0] and len(r) > ki: kind[(e, str(r[0]).strip())] = r[ki]
    skipped = []

    upload, internal, review = [], [], []
    seq, seqmap = 0, {}
    for l in sorted(lines, key=lambda l: (l['date'], l['entity'], l['slip'])):
        if l['cust'] in INTERNAL[l['entity']]:
            internal.append(l); continue
        if PRICE_FIX.search(l['memo']):
            review.append((l, '단가 정정/차액 건 - 재고 이동 없음으로 보고 제외')); continue
        if not l['code']:
            review.append((l, '품목코드 없음')); continue
        if l['code'].upper().startswith('Z-'):
            review.append((l, '매출할인 등 비재고 품목 - 제외')); continue
        code = cmap.get((l['entity'], l['code']), l['code'])
        # 생산품만 재고 ERP에 기록: 회계 ERP 품목구분이 [제품]인 것
        if kind and kind.get((l['entity'], l['code'])) != '[제품]':
            skipped.append(l); continue
        if stock is not None and code not in stock:
            review.append((l, f'재고 ERP에 없는 품목코드: {code} (매핑표에 추가 필요)')); continue
        key = (l['entity'], l['date'], l['slip'])
        if key not in seqmap:
            seq += 1; seqmap[key] = seq
        memo = f"{l['entity']} {l['date']:%y/%m/%d}-{l['slip']}" + (f" | {l['memo']}" if l['memo'] else '')
        upload.append([l['date'].strftime('%Y%m%d'), seqmap[key], None, l['cust'], None, WAREHOUSE,
                       None, None, None, code, None, None, l['qty'], 0, None, 0, 0, memo[:200], None])

    wb = openpyxl.Workbook()
    H, HF = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor='305496')

    def sheet(ws, header, data, widths):
        ws.append(header)
        for c in ws[1]: c.font, c.fill = H, HF
        for d in data: ws.append(d)
        for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = 'A2'

    ws = wb.active; ws.title = '업로드용'
    sheet(ws, TEMPLATE, upload, [10, 6, 10, 28, 8, 10, 8, 6, 6, 14, 8, 8, 8, 6, 8, 8, 8, 50, 8])
    for r in ws.iter_rows(min_row=2, min_col=1, max_col=1):
        for c in r: c.number_format = '@'
    row = lambda l: [l['entity'], f"{l['date']:%Y-%m-%d}-{l['slip']}", l['cust'], l['code'], l['item'], l['qty'], l['memo']]
    hdr = ['법인', '전표', '거래처', '품목코드', '품명 및 규격', '수량', '비고']
    sheet(wb.create_sheet('확인필요'), ['사유'] + hdr, [[why] + row(l) for l, why in review], [40, 6, 14, 24, 14, 40, 8, 30])
    sheet(wb.create_sheet('제외_내부거래'), hdr, [row(l) for l in internal], [6, 14, 14, 14, 40, 8, 30])
    s = wb.create_sheet('요약')
    for e in ('남경', '이엔지'):
        n = [l for l in lines if l['entity'] == e]
        s.append([e, '판매 줄', len(n)])
    s.append(['업로드 줄', len(upload)]); s.append(['업로드 전표', seq])
    s.append(['상품(매입품) 제외', len(skipped)])
    s.append(['내부거래 제외', len(internal)]); s.append(['확인필요', len(review)])
    wb.save(a.out)
    print(f'업로드 {len(upload)}줄/{seq}전표, 상품 제외 {len(skipped)}, 내부거래 제외 {len(internal)}, 확인필요 {len(review)} → {a.out}')


if __name__ == '__main__':
    main()
