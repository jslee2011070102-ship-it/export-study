#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""수출 거래 JSON → Commercial Invoice + Packing List (xlsx).

상업송장을 단일 원천으로 두고 포장명세서를 파생시킨다.
forms/수출서류_상업송장_포장명세서.xlsx 와 같은 레이아웃을 코드로 생성한 것으로,
엑셀을 손으로 채우는 대신 시스템이 가진 거래 데이터에서 바로 뽑을 때 쓴다.

    pip install openpyxl
    python3 automation/generate.py automation/samples/sample.json -o out/
"""
import argparse, json, os, re, sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill

A = 'Arial'
BLACK = Font(name=A, size=10)
BOLD  = Font(name=A, size=10, bold=True)
TITLE = Font(name=A, size=16, bold=True)
HEAD  = Font(name=A, size=10, bold=True, color='FFFFFF')
NOTE  = Font(name=A, size=9, italic=True, color='808080')
GRAY  = PatternFill('solid', fgColor='404040')
LIGHT = PatternFill('solid', fgColor='F2F2F2')
_t    = __import__('openpyxl').styles.Side(style='thin', color='999999')
BOX   = Border(left=_t, right=_t, top=_t, bottom=_t)
CTR   = Alignment(horizontal='center', vertical='center')

WIN_BAD = r'[<>:"/\\|?*\x00-\x1f]'


def safe(s):
    return re.sub(r'\s+', ' ', re.sub(WIN_BAD, ' ', s)).strip()[:80] or 'document'


def party_lines(p):
    return [p.get('상호', ''), p.get('주소1', ''), p.get('주소2', ''), p.get('담당자', '')]


def build(ws, d, is_pl):
    거래, 품목 = d.get('거래', {}), d.get('품목', [])
    ws['A1'] = 'PACKING LIST' if is_pl else 'COMMERCIAL INVOICE'
    ws['A1'].font = TITLE
    ws.merge_cells('A1:H1')
    ws['A1'].alignment = CTR

    r = 3
    ws.cell(r, 1, 'Shipper / Exporter').font = BOLD
    ws.cell(r, 5, 'Invoice No.').font = BOLD
    ws.cell(r, 6, 거래.get('invoice_no', '')).font = BLACK
    ws.cell(r, 7, 'Date').font = BOLD
    ws.cell(r, 8, 거래.get('invoice_date', '')).font = BLACK
    for i, line in enumerate(party_lines(d.get('수출자', {}))):
        ws.cell(r + 1 + i, 1, line).font = BLACK
    ws.cell(r + 1, 5, 'P/O No.').font = BOLD
    ws.cell(r + 1, 6, 거래.get('po_no', '')).font = BLACK

    r = 8
    ws.cell(r, 1, 'Consignee').font = BOLD
    for i, line in enumerate(party_lines(d.get('수입자', {}))):
        ws.cell(r + 1 + i, 1, line).font = BLACK
    ws.cell(r, 5, 'Notify Party').font = BOLD
    ws.cell(r + 1, 5, d.get('착화통지처', 'SAME AS CONSIGNEE')).font = BLACK

    r = 13
    pairs = [('Port of Loading', 거래.get('선적항', '')),
             ('Port of Discharge', 거래.get('도착항', '')),
             ('Final Destination', 거래.get('최종목적지', '')),
             ('Vessel / Voyage', 거래.get('선박명', 'T.B.A.')),
             ('ETD', 거래.get('출항예정일', 'T.B.A.')),
             ('Country of Origin', 거래.get('원산지', '')),
             ('Price Term', 거래.get('가격조건', ''))]
    if not is_pl:
        pairs += [('Incoterms', 거래.get('인코텀즈버전', 'Incoterms(R) 2020')),
                  ('Payment Term', 거래.get('결제조건', '')),
                  ('Currency', 거래.get('통화', ''))]
    for i, (lab, val) in enumerate(pairs):
        row, col = r + i // 2, 1 if i % 2 == 0 else 5
        ws.cell(row, col, lab).font = BOLD
        ws.cell(row, col + 1, val).font = BLACK

    tr = r + (len(pairs) + 1) // 2 + 1
    if is_pl:
        hdr = ['No.', 'Description', 'Spec', 'HS CODE', "Q'ty", 'Unit', 'N.W.(kg)', 'G.W.(kg)', 'Packages']
        keys = ['품명', '규격', 'hs_code', '수량', '단위', '순중량_kg', '총중량_kg', '포장수량']
    else:
        hdr = ['No.', 'Description', 'Spec', 'HS CODE', "Q'ty", 'Unit', 'Unit Price', 'Amount']
        keys = ['품명', '규격', 'hs_code', '수량', '단위', '단가', None]
    for i, h in enumerate(hdr, 1):
        c = ws.cell(tr, i, h)
        c.font, c.fill, c.border, c.alignment = HEAD, GRAY, BOX, CTR

    first = tr + 1
    for n, it in enumerate(품목):
        row = first + n
        ws.cell(row, 1, n + 1).font = BLACK
        ws.cell(row, 1).border = BOX
        for i, k in enumerate(keys, 2):
            c = ws.cell(row, i)
            c.border = BLACK and BOX
            c.font = BLACK
            if k is None:                       # Amount = 수량 x 단가 (수식으로)
                c.value = f'=IF(E{row}="","",E{row}*G{row})'
            else:
                c.value = it.get(k)
    last = first + len(품목) - 1

    tot = last + 1
    ws.cell(tot, 1, 'TOTAL').font = BOLD
    for i in range(1, len(hdr) + 1):
        ws.cell(tot, i).fill, ws.cell(tot, i).border = LIGHT, BOX
    sum_cols = {'E'}                            # 수량
    sum_cols |= {'G', 'H', 'I'} if is_pl else {'H'}
    for i, h in enumerate(hdr, 1):
        letter = chr(64 + i)
        if letter in sum_cols:
            c = ws.cell(tot, i, f'=SUM({letter}{first}:{letter}{last})')
            c.font, c.fill, c.border = BOLD, LIGHT, BOX

    er = tot + 2
    if is_pl:
        pk = d.get('포장', {})
        ws.cell(er, 1, 'Total Packages').font = BOLD
        ws.cell(er, 2, f'=I{tot}&" {pk.get("포장단위", "CTNS")}"').font = BLACK
        ws.cell(er + 1, 1, 'Total Measurement (CBM)').font = BOLD
        ws.cell(er + 1, 2, pk.get('총부피_CBM')).font = BLACK
    else:
        ws.cell(er, 1, 'Total Amount').font = BOLD
        ws.cell(er, 2, f'="{거래.get("통화", "")} "&TEXT(H{tot},"#,##0.00")').font = BOLD
        ws.cell(er + 1, 1, 'Price Term').font = BOLD
        ws.cell(er + 1, 2, f'{거래.get("가격조건", "")} {거래.get("인코텀즈버전", "")}'.strip()).font = BLACK

    sr = er + 3
    ws.cell(sr, 6, d.get('수출자', {}).get('상호', '')).font = BOLD
    ws.cell(sr + 1, 6, 'Signed by').font = NOTE
    ws.cell(sr + 2, 6, '________________________').font = BLACK

    for col, w in zip('ABCDEFGHI', [8, 34, 18, 14, 10, 8, 14, 16, 12]):
        ws.column_dimensions[col].width = w


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('file')
    ap.add_argument('-o', '--out', default='out')
    a = ap.parse_args()

    with open(a.file, encoding='utf-8') as f:
        d = json.load(f)
    os.makedirs(a.out, exist_ok=True)
    inv = d.get('거래', {}).get('invoice_no', 'DOC')

    wb = Workbook()
    build(wb.active, d, is_pl=False)
    wb.active.title = 'Commercial Invoice'
    ci = os.path.join(a.out, f'{safe(inv)}_CI.xlsx')
    wb.save(ci)

    wb = Workbook()
    build(wb.active, d, is_pl=True)
    wb.active.title = 'Packing List'
    pl = os.path.join(a.out, f'{safe(inv)}_PL.xlsx')
    wb.save(pl)

    print(f'생성: {ci}\n생성: {pl}')
    print('\n주의 - openpyxl 은 수식의 계산값을 쓰지 않는다.')
    print('      엑셀에서 열면 자동 계산되지만, 프로그램으로 값을 읽으려면 recalc 가 필요하다.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
