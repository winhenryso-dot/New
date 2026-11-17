"""회계 ERP(법인 2곳) 품목코드 → 재고 ERP 품목코드 매핑표 초안 생성.

근거 우선순위
 1. 판매이력: 같은 날짜·수량으로 회계 ERP와 재고 ERP에 각각 입력된 판매 → 실제로 쓰던 짝
 2. 이름일치: 품목명[규격]이 공백·쉼표 무시하고 같음
 3. 후보(확인필요): 가장 비슷한 이름 (재고ERP품목코드 칸은 비워 두고 후보 칸에만 표시)

사용법:
  python3 build_item_map.py --stock-items 재고ERP_품목등록.xlsx \
     --items-namkyung 남경_품목등록.xlsx --items-eng 이엔지_품목등록.xlsx \
     [--sales-namkyung 남경_판매현황.xlsx --sales-eng 이엔지_판매현황.xlsx --stock-sales 재고ERP_판매현황.xlsx] \
     -o 품목코드_매핑표.xlsx
"""
import argparse, re, sys, os, collections, difflib
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ecount_xlsx import read_rows
import sales_to_stock as sts

norm = lambda t: re.sub(r'[\s,]', '', str(t or '')).upper()


def table(path, first='품목코드'):
    rows = read_rows(path)
    hi = next(i for i, r in enumerate(rows) if r and r[0] == first)
    h = [str(x or '').strip() for x in rows[hi]]
    out = []
    for r in rows[hi + 1:]:
        if r and r[0] and len(r) > 1 and r[1] is not None:
            out.append(dict(zip(h, r + [None] * (len(h) - len(r)))))
    return out


def full(name, spec):
    return f'{name} [{spec}]' if spec else str(name)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--stock-items', required=True)
    ap.add_argument('--items-namkyung', required=True); ap.add_argument('--items-eng', required=True)
    ap.add_argument('--sales-namkyung'); ap.add_argument('--sales-eng'); ap.add_argument('--stock-sales')
    ap.add_argument('-o', '--out', required=True)
    a = ap.parse_args()

    SI = {str(r['품목코드']).strip(): r for r in table(a.stock_items)}
    sname = {c: r.get('품목명[규격]') or full(r['품목명'], r.get('규격정보')) for c, r in SI.items()}
    by_name = collections.defaultdict(list)
    for c, n in sname.items(): by_name[norm(n)].append(c)
    snames = list(by_name)

    acct = {}
    for e, f in (('남경', a.items_namkyung), ('이엔지', a.items_eng)):
        for r in table(f):
            if r.get('품목구분') == '[제품]':
                acct[(e, str(r['품목코드']).strip())] = full(r['품목명'], r.get('규격명'))

    # 1) 판매이력으로 짝 찾기
    hist = collections.defaultdict(collections.Counter)
    if a.stock_sales and (a.sales_namkyung or a.sales_eng):
        st = collections.defaultdict(list)
        rows = read_rows(a.stock_sales)
        for r in rows:
            m = re.match(r'(\d{4})/(\d\d)/(\d\d)\s*-(\d+)', str(r[0] if r else ''))
            if m and len(r) > 2:
                codes = by_name.get(norm(r[1]))
                if codes: st[(f'{m[1]}{m[2]}{m[3]}', r[2])].append((norm(r[1]), codes[0]))
        L = []
        if a.sales_namkyung: L += sts.load_sales(a.sales_namkyung, '남경')
        if a.sales_eng: L += sts.load_sales(a.sales_eng, '이엔지')
        for l in L:
            k = (l['date'].strftime('%Y%m%d'), l['qty'])
            if (l['entity'], l['code']) not in acct or not st.get(k):
                continue
            # 같은 날짜·수량 후보 중 이름이 가장 비슷한 것
            an = norm(l['item'])
            best = max(st[k], key=lambda x: difflib.SequenceMatcher(None, an, x[0]).ratio())
            if difflib.SequenceMatcher(None, an, best[0]).ratio() >= 0.6:
                hist[(l['entity'], l['code'])][best[1]] += 1

    out, stat = [], collections.Counter()
    for (e, c), n in sorted(acct.items()):
        code, why, cand, score = None, '', None, None
        if hist.get((e, c)):
            code, cnt = hist[(e, c)].most_common(1)[0]; why = f'판매이력 {cnt}건 일치'
        elif by_name.get(norm(n)):
            code = by_name[norm(n)][0]; why = '품목명[규격] 일치'
            if c in SI and c != code: why += f' (같은 코드 {c}는 다른 품목)'
        else:
            m = difflib.get_close_matches(norm(n), snames, n=1, cutoff=0.75)
            if m:
                cand = by_name[m[0]][0]; score = round(difflib.SequenceMatcher(None, norm(n), m[0]).ratio(), 2)
                why = '확인필요: 비슷한 이름 후보'
            else:
                why = '확인필요: 재고 ERP에서 못 찾음'
        stat[why.split(' ')[0].split(':')[0]] += 1
        out.append([e, c, code, why, n, sname.get(code or cand), cand, score])

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = '매핑표'
    hdr = ['법인', '법인품목코드', '재고ERP품목코드', '근거', '회계 품목명[규격]', '재고 품목명[규격]', '후보코드', '유사도']
    ws.append(hdr)
    H, HF = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor='305496')
    Y = PatternFill('solid', fgColor='FFF2CC')
    for c in ws[1]: c.font, c.fill = H, HF
    for r in out:
        ws.append(r)
        if r[3].startswith('확인필요'):
            for c in ws[ws.max_row]: c.fill = Y
    for i, w in enumerate([6, 14, 14, 30, 45, 45, 14, 7], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions
    wb.save(a.out)
    print(dict(stat), '→', a.out)


if __name__ == '__main__':
    main()
