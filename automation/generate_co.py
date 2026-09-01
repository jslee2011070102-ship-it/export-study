#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""거래 데이터(JSON) → 원산지증명서 / 원산지신고서 문안.

협정마다 성격이 다르다. FTA특례법 시행규칙 제15조 기준.

    기관발급   세관(유니패스)/상공회의소가 발급. 수출자가 만들 수 없으므로 안내만 한다
    서식형     정해진 별지서식을 수출자가 작성. 한-미만 구현
    문안형     별도 서식 없이 상업송장 등에 규정된 문안을 기재. 한-영/한-EU 구현

    python3 automation/generate_co.py automation/samples/sample.json -o out/
    python3 automation/generate_co.py <file.json> -o out/ --협정 한-영국
    python3 automation/generate_co.py <file.json> --협정 한-베트남      # 안내만 출력

산출물 (협정에 따라)
    {invoice}_원산지신고서_{협정}.txt / .pdf    문안형. 상업송장에 붙여넣을 문안
    {invoice}_CO_US.xlsx / .pdf                 한-미 원산지증명서
"""
import argparse, datetime as dt, json, os, re, sys

import openpyxl
from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import co_rules as R
from build_template import 셀, 인쇄설정, A, 가운데, 왼쪽, 오른쪽, 줄바꿈, 회색, 노랑

가운데줄바꿈 = A(horizontal='center', vertical='center', wrap_text=True)
from generate import 통째로PDF

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
품목_최대 = 15


def _금(v):
    return v if v not in (None, '') else ''


# ══════════════════════════════════════════════════════════════
# 문안형 — 상업서류에 기재할 원산지신고서 문안
# ══════════════════════════════════════════════════════════════
def 문안생성(d, 협정키, outdir, invoice):
    co = d.get('원산지증명', {}) or {}
    수출자 = d.get('수출자', {}) or {}
    인증번호 = (co.get('인증수출자번호') or '').strip()
    원산지 = co.get('원산지영문') or 'the Republic of Korea'
    장소 = co.get('작성장소') or (수출자.get('주소2') or 수출자.get('주소1') or '')
    작성일 = co.get('작성일') or d.get('거래', {}).get('invoice_date') or dt.date.today().isoformat()
    서명자 = co.get('서명자') or 수출자.get('담당자') or ''

    영문, 확정 = R.문안만들기(협정키, 인증번호, 원산지, 국문=False)
    국문, _ = R.문안만들기(협정키, 인증번호, '대한민국', 국문=True)
    if 영문 is None:
        return None

    정보 = R.협정[협정키]
    줄 = []
    줄.append(f'원산지신고서 (Origin Declaration) — 한-{협정키} FTA')
    줄.append('=' * 78)
    줄.append(f'근거   FTA특례법 시행규칙 {정보["근거"]} / {정보.get("문안근거","")}')
    줄.append(f'방식   별도 서식 없음. 아래 문안을 상업송장 등 상업서류에 기재')
    if not 확정:
        줄.append('')
        줄.append('!! [검증] 이 문안은 원문 대조를 마치지 않았다.')
        줄.append(f'   {정보.get("주의","")}')
        줄.append('   구조가 같은 한-영국(별표 20의2) 문안을 근거로 작성한 것이므로,')
        줄.append('   사용 전 관세청 FTA포털 또는 대한상공회의소에서 현행 문안을 확인할 것.')
    줄.append('')
    줄.append('── 영문 ' + '─' * 70)
    줄.append('')
    줄.append(영문)
    줄.append('')
    줄.append(f'{장소}, {작성일}')
    줄.append('')
    줄.append(f'{서명자}')
    줄.append('(Signature of the exporter, in addition the name of the person signing '
              'the declaration has to be indicated in clear script)')
    줄.append('')
    줄.append('── 국문 ' + '─' * 70)
    줄.append('')
    줄.append(국문)
    줄.append('')
    줄.append('── 작성방법 ' + '─' * 66)
    for t in R.영국_작성방법:
        줄.append('  ' + t)
    줄.append('')
    줄.append('── 확인사항 ' + '─' * 66)
    if not 인증번호:
        줄.append('  · 인증수출자번호가 비어 있다. 한-EU/한-영국은 건당 6,000유로 초과 시')
        줄.append('    인증수출자만 원산지신고서를 작성할 수 있다 (시행규칙 제7조제2항).')
    else:
        줄.append(f'  · 인증수출자번호 {인증번호}')
    줄.append('  · 이 문안은 상업송장 자체에 기재해야 한다. 별지에 따로 만들면 안 된다.')
    줄.append('  · 상호/주소는 통관고유부호 등록정보와 대소문자까지 일치시킬 것.')

    본문 = '\n'.join(줄) + '\n'
    dst = os.path.join(outdir, f'{invoice}_원산지신고서_{협정키}.txt')
    with open(dst, 'w', encoding='utf-8') as f:
        f.write(본문)
    return dst, 영문, 국문, 확정


# ══════════════════════════════════════════════════════════════
# 서식형 — 한-미 FTA 원산지증명서 (별지 제17호서식)
# ══════════════════════════════════════════════════════════════
def 미국CO(d, outdir, invoice):
    거래 = d.get('거래', {}) or {}
    co = d.get('원산지증명', {}) or {}
    수출자 = d.get('수출자', {}) or {}
    수입자 = d.get('수입자', {}) or {}
    품목 = d.get('품목', [])
    생산자 = co.get('생산자') or {}
    if isinstance(생산자, str):
        생산자 = {'상호': 생산자}

    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Certificate of Origin')
    for col, w in zip('ABCDEFGH', [13, 24, 13, 24, 11, 11, 11, 12]):
        ws.column_dimensions[col].width = w

    셀(ws, 'A1', 'Certificate of Origin', bold=True, size=15, align=가운데, border=False, merge='A1:H1')
    셀(ws, 'A2', 'Korea-US Free Trade Agreement', bold=True, size=11, align=가운데,
      border=False, merge='A2:H2')
    ws.row_dimensions[1].height = 22

    def 블록(row, 번호, 제목, 정보, 열='A'):
        끝 = 'D' if 열 == 'A' else 'H'
        중 = 'B' if 열 == 'A' else 'F'
        셀(ws, f'{열}{row}', f'{번호}. {제목}', bold=True, fill=회색,
          merge=f'{열}{row}:{끝}{row}')
        for i, (라, 값) in enumerate(정보, start=1):
            셀(ws, f'{열}{row+i}', 라, fill=회색)
            셀(ws, f'{중}{row+i}', 값, align=줄바꿈, merge=f'{중}{row+i}:{끝}{row+i}')
        return row + len(정보) + 1

    def 당사자(p):
        return [('Name (성명)', p.get('상호') or ''),
                ('Address (주소)', ' '.join(x for x in [p.get('주소1'), p.get('주소2')] if x)),
                ('Contact (담당자/전화)', p.get('담당자') or ''),
                ('E-mail (전자주소)', p.get('email') or '')]

    r = 4
    끝1 = 블록(r, '1', 'Exporter (수출자)', 당사자(수출자), 'A')
    포괄 = co.get('포괄증명기간') or ''
    셀(ws, 'E4', '2. Blanket Period (원산지포괄증명기간)', bold=True, fill=회색, merge='E4:H4')
    셀(ws, 'E5', 'From (부터)', fill=회색); 셀(ws, 'F5', 포괄.split('~')[0].strip() if 포괄 else '',
                                            merge='F5:H5')
    셀(ws, 'E6', 'To (까지)', fill=회색); 셀(ws, 'F6', 포괄.split('~')[-1].strip() if '~' in 포괄 else '',
                                          merge='F6:H6')
    for rr in (7, 8):
        셀(ws, f'E{rr}', None, fill=None, merge=f'E{rr}:H{rr}')

    r = 끝1 + 1
    끝2 = 블록(r, '3', 'Producer (생산자)', 당사자(생산자), 'A')
    블록(r, '4', 'Importer (수입자)', 당사자(수입자), 'E')

    r = max(끝2, r + 5) + 1
    헤더 = ['5. Description of Good(s)', '', '6. HS No.(6)', '7. Preference Criterion',
            '8. Producer', '9. Country of Origin']
    셀(ws, f'A{r}', '5. Description of Good(s)', bold=True, align=가운데, fill=회색, merge=f'A{r}:C{r}')
    셀(ws, f'D{r}', '6. HS No.\n(6단위)', bold=True, align=가운데줄바꿈, fill=회색)
    셀(ws, f'E{r}', '7. Preference\nCriterion', bold=True, align=가운데줄바꿈, fill=회색)
    셀(ws, f'F{r}', '8. Producer', bold=True, align=가운데, fill=회색)
    셀(ws, f'G{r}', '9. Country of\nOrigin', bold=True, align=가운데줄바꿈, fill=회색, merge=f'G{r}:H{r}')
    ws.row_dimensions[r].height = 30

    기준 = co.get('원산지결정기준코드') or ''
    국가 = co.get('원산지약어') or 'KR'
    생산자여부 = 'YES' if (생산자.get('상호') and
                          생산자.get('상호') == 수출자.get('상호')) else 'NO'
    for i in range(품목_최대):
        rr = r + 1 + i
        it = 품목[i] if i < len(품목) else None
        셀(ws, f'A{rr}', (it.get('품명') if it else None), align=줄바꿈, merge=f'A{rr}:C{rr}')
        hs = (it.get('hs_code') or '') if it else ''
        셀(ws, f'D{rr}', re.sub(r'\D', '', hs)[:6] if hs else None, align=가운데)
        셀(ws, f'E{rr}', 기준 if it else None, align=가운데)
        셀(ws, f'F{rr}', 생산자여부 if it else None, align=가운데)
        셀(ws, f'G{rr}', 국가 if it else None, align=가운데, merge=f'G{rr}:H{rr}')
    r = r + 1 + 품목_최대

    셀(ws, f'A{r}', '10. Remarks', bold=True, fill=회색, merge=f'A{r}:H{r}')
    셀(ws, f'A{r+1}', co.get('비고') or '', align=줄바꿈, merge=f'A{r+1}:H{r+2}')
    ws.row_dimensions[r + 1].height = 18

    r += 3
    확인문 = ('I certify that the information on this document is true and accurate and I assume '
             'the responsibility for proving such representations. I understand that I am liable '
             'for any false statements or material omissions made on or in connection with this '
             'document. I agree to maintain and present upon request, documentation necessary to '
             'support this certificate, and to inform, in writing, all persons to whom the '
             'certificate was given of any changes that would affect the accuracy or validity of '
             'this certificate.')
    셀(ws, f'A{r}', 확인문, align=줄바꿈, size=9, merge=f'A{r}:H{r+3}')
    ws.row_dimensions[r].height = 52

    r += 4
    서명 = [('11. Authorized Signature', ''),
            ('12. Company', 수출자.get('상호') or ''),
            ('13. Name (성명)', co.get('서명자') or 수출자.get('담당자') or ''),
            ('14. Title (직위)', co.get('직위') or ''),
            ('15. Date (작성일)', co.get('작성일') or 거래.get('invoice_date') or ''),
            ('16. Contact (담당자/전화)', 수출자.get('담당자') or '')]
    for i, (라, 값) in enumerate(서명):
        rr = r + i
        셀(ws, f'A{rr}', 라, bold=True, fill=회색, merge=f'A{rr}:B{rr}')
        셀(ws, f'C{rr}', 값, merge=f'C{rr}:H{rr}')
    끝행 = r + len(서명)

    인쇄설정(ws, f'A1:H{끝행}')
    dst = os.path.join(outdir, f'{invoice}_CO_US.xlsx')
    wb.save(dst)
    return dst


# ══════════════════════════════════════════════════════════════
def 안내(협정키):
    정보 = R.협정[협정키]
    방식 = 정보['방식']
    print(f'\n협정   한-{협정키} FTA')
    print(f'방식   {방식}    (FTA특례법 시행규칙 {정보["근거"]})')
    if 방식 == '기관발급':
        print(f'서식   {정보.get("서식","")}')
        print('\n  이 협정의 원산지증명서는 세관 또는 상공회의소가 발급한다.')
        print('  수출자가 직접 만든 문서는 효력이 없으므로 생성하지 않는다.')
        print('\n  발급 경로')
        print('    · 세관  — 유니패스 (unipass.customs.go.kr). 수수료 없음')
        print('    · 상의  — 대한상공회의소 원산지증명센터 (cert.korcham.net). 비회원 수수료')
        print('\n  발급 신청 시 필요한 서류 (forms/공식서식/ 에 원본 보관)')
        print('    · 원산지증명서 발급(재발급/정정발급)신청서')
        print('    · 원산지소명서')
        print('    · 원산지(포괄)확인서 — 원재료가 국내산임을 공급자가 증명')
        print('    · 국내제조(포괄)확인서')
        print('    · 상업송장 / 포장명세서 / BOM / 제조공정도')
        print('\n  먼저 automation/validate.py 로 CO 규칙(발급형식/미소기준/PSR)을 확인할 것.')
    elif 방식 == '서식형':
        print(f'서식   {정보.get("서식","")}')
        print('\n  수출자가 직접 작성하는 서식이나 이 도구는 한-미만 구현했다.')
        print(f'  forms/공식서식/ 또는 국가법령정보센터에서 {정보.get("서식","")} 을 받아 작성할 것.')
    elif 방식 == '문안형':
        print(f'문안근거  {정보.get("문안근거","")}')
        print('\n  별도 서식 없이 상업송장 등에 문안을 기재하는 방식이나,')
        print('  해당 별표 원문을 확보하지 못해 문안을 생성하지 않는다.')
        print('  관세청 FTA포털 또는 대한상공회의소에서 현행 문안을 확인할 것.')
    else:
        print('\n  발급 경로가 둘 이상이다.')
        if 정보.get('비고'):
            print(f'  {정보["비고"]}')
        print('  기관발급분은 세관/상의, 자율발급분은 인증수출자 자격이 필요한 경우가 많다.')
    if 정보.get('비고') and 방식 != '혼합':
        print(f'\n  비고  {정보["비고"]}')


def main():
    ap = argparse.ArgumentParser(description='거래 데이터 → 원산지증명서 / 원산지신고서 문안')
    ap.add_argument('데이터', help='거래 마스터 JSON')
    ap.add_argument('-o', '--out', default='.', help='출력 폴더')
    ap.add_argument('--협정', help='협정명. 없으면 JSON 의 원산지증명.적용협정 을 쓴다')
    ap.add_argument('--no-pdf', action='store_true')
    ap.add_argument('--timeout', type=int, default=300)
    a = ap.parse_args()

    with open(a.데이터, encoding='utf-8') as f:
        d = json.load(f)
    이름 = a.협정 or (d.get('원산지증명', {}) or {}).get('적용협정')
    협정키 = R.찾기(이름)
    if not 협정키:
        print(f'협정을 알 수 없다: "{이름}"')
        print('사용 가능: ' + ' / '.join(sorted(R.협정)))
        return 1

    정보 = R.협정[협정키]
    if 정보['방식'] == '기관발급' or (정보['방식'] == '문안형' and 정보.get('문안') is None) \
            or (정보['방식'] == '서식형' and 협정키 != '미국') or 정보['방식'] == '혼합':
        안내(협정키)
        return 0

    invoice = re.sub(r'[^\w.-]', '_', (d.get('거래', {}).get('invoice_no') or 'INVOICE'))
    os.makedirs(a.out, exist_ok=True)

    if 정보['방식'] == '문안형':
        결과 = 문안생성(d, 협정키, a.out, invoice)
        if not 결과:
            안내(협정키)
            return 0
        dst, 영문, 국문, 확정 = 결과
        print(f'\n협정   한-{협정키} FTA   ({정보["방식"]}, 시행규칙 {정보["근거"]} / {정보.get("문안근거","")})')
        if not 확정:
            print('\n!! [검증] 문안 원문을 대조하지 못했다. 사용 전 관세청 FTA포털 확인 필요.')
            print(f'   {정보.get("주의","")}')
        print(f'\n생성: {dst}')
        print('\n── 상업송장에 기재할 문안 ' + '─' * 44)
        print(영문)
        return 0

    # 한-미
    xlsx = 미국CO(d, a.out, invoice)
    print(f'\n협정   한-미국 FTA   (서식형, 시행규칙 제15조제8항)')
    print(f'생성: {xlsx}')
    print('       별지 제17호서식 기준. 제15조제8항제2호의 8개 기재사항을 모두 담고 있다.')
    if not a.no_pdf:
        pdf = os.path.join(a.out, f'{invoice}_CO_US.pdf')
        n = 통째로PDF(xlsx, pdf, a.timeout)
        if n:
            print(f'생성: {pdf}  ({n}장)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
