#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""수출 거래 데이터 정합성 검증.

manual/부록C_실무함정_10선.md 과 법령 검증 결과(docs/11, docs/12)를 규칙으로 옮긴 것.
서류를 만들기 전에 돌려서 사고를 미리 잡는 용도.

    python3 automation/validate.py automation/samples/sample.json
    python3 automation/validate.py <file.json> --today 2026-09-20
"""
import argparse, csv, datetime as dt, json, os, re, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PSR  = os.path.join(ROOT, 'data', 'fta_psr.csv')
공휴일표 = os.path.join(ROOT, 'data', '공휴일.csv')

# ── 검증 완료된 법령 상수 (docs/11, docs/12) ────────────────────────
적재기한일수 = 30                      # 관세법 제251조제1항
목록통관_FOB_한도_원 = 2_000_000        # 수출 및 반송통관에 관한 고시 제55조제1항제8호
인증수출자_기준_EUR = 6000              # FTA특례법 시행규칙 제7조제2항 (EU/영국)

발급형식 = {                            # FTA특례법 시행규칙 제7조
    '기관발급만': ['싱가포르', '아세안', '인도', '베트남', '중국', '인도네시아'],
    '자율발급만': ['칠레', 'EU', '페루', '미국', '튀르키예', '콜롬비아',
                 '캐나다', '뉴질랜드', '중미', '영국'],
    '병존':      ['호주', '이스라엘', 'RCEP', '캄보디아', '필리핀', 'UAE', '에콰도르'],
}
미소기준 = {                            # FTA특례법 시행규칙 별표 (확인된 6개 협정)
    '아세안': (10, 'FOB'), '베트남': (10, 'FOB'), 'RCEP': (10, 'FOB'),
    '필리핀': (10, 'FOB'), 'UAE': (15, '공장도가격'), '에콰도르': (10, 'FOB'),
}
모호한_품명 = re.compile(r'^\s*(PARTS?|GOODS?|ITEMS?|SAMPLES?|ACCESSORIES|MATERIALS?|기타|부품)\s*$', re.I)
운임필요_조건 = ('CFR', 'CPT', 'CIF', 'CIP')
보험필요_조건 = ('CIF', 'CIP')
해상전용_조건 = ('FAS', 'FOB', 'CFR', 'CIF')   # 인도지점이 선측/본선이라 해상·내수로 전용
해상전용_대체 = {'FAS': 'FCA', 'FOB': 'FCA', 'CFR': 'CPT', 'CIF': 'CIP'}
C그룹_조건   = ('CPT', 'CIP', 'CFR', 'CIF')    # 위험분기점과 비용분기점이 다름
D그룹_조건   = ('DAP', 'DPU', 'DDP')           # 목적지 도착까지 매도인이 위험 부담
인코텀즈_11 = ('EXW', 'FCA', 'FAS', 'FOB', 'CPT', 'CIP', 'CFR', 'CIF',
              'DAP', 'DPU', 'DDP')


def _공휴일():
    """관세법 제8조제3항 - 기한이 토/일/공휴일/대체공휴일/노동절이면 다음 날이 기한."""
    s = set()
    if os.path.exists(공휴일표):
        with open(공휴일표, encoding='utf-8') as f:
            for row in csv.DictReader(f):
                s.add(dt.date.fromisoformat(row['일자']))
    return s

공휴일 = _공휴일()
공휴일_수록연도 = {d.year for d in 공휴일}


def 적재기한(수리일):
    """수리일 + 30일. 말일이 휴일이면 관세법 제8조제3항에 따라 다음 개청일로 밀린다."""
    d = 수리일 + dt.timedelta(days=적재기한일수)
    밀림 = 0
    while d.weekday() >= 5 or d in 공휴일:
        d += dt.timedelta(days=1)
        밀림 += 1
    return d, 밀림


class Report:
    def __init__(self):
        self.rows = []

    def add(self, level, code, msg, fix=''):
        self.rows.append({'level': level, 'code': code, 'msg': msg, 'fix': fix})

    err  = lambda self, c, m, f='': self.add('ERROR', c, m, f)
    warn = lambda self, c, m, f='': self.add('WARN',  c, m, f)
    info = lambda self, c, m, f='': self.add('INFO',  c, m, f)

    @property
    def errors(self):
        return [r for r in self.rows if r['level'] == 'ERROR']


def load_psr():
    if not os.path.exists(PSR):
        return []
    with open(PSR, encoding='utf-8-sig') as f:
        return list(csv.DictReader(f))


def find_psr(rows, hs, 협정키):
    """HS 코드로 품목별 원산지결정기준을 찾는다. 6단위 → 4단위 순으로 넓혀가며 조회."""
    if not hs:
        return None
    digits = hs.replace('.', '')
    cands = [hs, f'{digits[:4]}.{digits[4:6]}', f'{digits[:2]}.{digits[2:4]}', digits[:4], digits[:2]]
    for c in cands:
        for r in rows:
            if 협정키 in r['협정'] and r['적용단위'] == c:
                return r
    return None


def validate(d, today):
    rep = Report()
    거래 = d.get('거래', {})
    품목 = d.get('품목', [])
    통관 = d.get('통관', {})
    co   = d.get('원산지증명', {})

    # ── 1. 인코텀즈 (02장)
    가격조건 = (거래.get('가격조건') or '').strip()
    if not 가격조건:
        rep.err('INCO-01', '가격조건(인코텀즈)이 없다.', '예) FOB BUSAN')
    else:
        규칙 = 가격조건.split()[0].upper()
        if len(가격조건.split()) < 2:
            rep.err('INCO-02', f'가격조건에 지정장소가 없다: "{가격조건}"',
                    '규칙만 쓰면 인도장소가 특정되지 않는다. 예) FOB BUSAN')
        if 규칙 in 운임필요_조건 and 거래.get('운임') is None:
            rep.err('INCO-03', f'{규칙} 조건인데 운임 금액이 없다.',
                    '수출신고 시 FOB 환산에 필요하다. 관세사에게 전달할 값이다.')
        if 규칙 in 보험필요_조건 and 거래.get('보험료') is None:
            rep.err('INCO-04', f'{규칙} 조건인데 보험료 금액이 없다.',
                    f'{규칙} 는 매도인의 부보 의무가 있다. 보험증권도 함께 준비할 것.')
        if 규칙 == 'EXW':
            rep.warn('INCO-05', 'EXW 는 수출통관도 매수인 의무다.',
                     '실무상 매도인이 대행하게 되므로 누가 할지 문서로 정할 것. FCA 가 더 깔끔하다.')
        if not re.search(r'2020', 거래.get('인코텀즈버전') or ''):
            rep.warn('INCO-06', '인코텀즈 버전이 2020 으로 명시되지 않았다.',
                     '계약서와 송장에 Incoterms(R) 2020 을 표기할 것.')
        if 규칙 not in 인코텀즈_11:
            rep.err('INCO-07', f'인코텀즈 2020 의 11개 규칙이 아니다: "{규칙}"',
                    'DAT 는 2020 에서 DPU 로 바뀌었다. 유효 규칙: ' + '/'.join(인코텀즈_11))
        운송방식 = (거래.get('운송방식') or '').strip()
        if 규칙 in 해상전용_조건 and 운송방식 in ('항공', '육상'):
            rep.err('INCO-08', f'{규칙} 는 해상/내수로 전용인데 운송방식이 {운송방식} 이다.',
                    f'인도지점이 선측/본선이라 성립하지 않는다. {해상전용_대체[규칙]} 로 바꿀 것.')
        if 규칙 in 해상전용_조건 and 거래.get('컨테이너화물') is True:
            rep.warn('INCO-09', f'{규칙} 를 컨테이너 화물에 쓰고 있다.',
                     f'컨테이너는 CY 에서 운송인에게 인도되는데 위험은 본선 적재까지 매도인이 진다. '
                     f'{해상전용_대체[규칙]} 가 정확하다.')
        if 규칙 in C그룹_조건 and not (거래.get('인도장소') or '').strip():
            rep.warn('INCO-10', f'{규칙} 인데 인도장소(위험분기점)가 지정되지 않았다.',
                     f'"{가격조건}" 의 지명은 비용분기점(목적지)이다. '
                     '위험이 넘어가는 국내 인도장소를 계약서에 따로 특정할 것.')
        if 규칙 in D그룹_조건:
            rep.warn('INCO-11', f'{규칙} 는 목적지 도착까지 매도인이 위험을 진다.',
                     '규칙상 부보 의무는 없으나 사고 시 전액 매도인 손해다. 운송보험을 들 것.')
        if 규칙 == 'DPU':
            rep.info('INCO-12', 'DPU 는 11개 중 유일하게 매도인 양하의무가 있다.',
                     '목적지에 지게차/크레인 등 하역 수단이 있는지 확인할 것. 양하 중 사고도 매도인 부담.')
        if 규칙 == 'DDP':
            rep.warn('INCO-13', 'DDP 는 수입국 관세/부가세까지 매도인 부담이다.',
                     '수입국에서 비거주자가 수입신고인이 될 수 있는지, 수입부가세 환급이 가능한지 '
                     '확인할 것. DAP + 관세 별도정산이 더 안전한 경우가 많다.')

    # ── 2. 품목 (06장)
    for i, it in enumerate(품목, 1):
        품명 = (it.get('품명') or '').strip()
        if not 품명 or 모호한_품명.match(품명):
            rep.err('ITEM-01', f'{i}번 품목의 품명이 모호하다: "{품명}"',
                    'HS 분류가 가능하도록 재질/용도/형태가 드러나게 쓸 것.')
        nw, gw = it.get('순중량_kg'), it.get('총중량_kg')
        if nw is not None and gw is not None and gw < nw:
            rep.err('ITEM-02', f'{i}번 품목의 총중량({gw})이 순중량({nw})보다 작다.',
                    '총중량은 포장 포함, 순중량은 포장 제외다.')
        if gw is None:
            rep.warn('ITEM-03', f'{i}번 품목의 총중량이 없다.',
                     '신고중량과 실제중량 차이는 미선적 처리의 주원인이다.')
        hs = it.get('hs_code')
        if not hs:
            rep.warn('ITEM-04', f'{i}번 품목의 HS 코드가 없다.', '관세사 분류 전이면 무시해도 된다.')

    # ── 3. 당사자 정합성 (06장 / 09장)
    for 역할 in ('수출자', '수입자'):
        p = d.get(역할, {})
        if not (p.get('상호') or '').strip():
            rep.err('PARTY-01', f'{역할} 상호가 없다.')
        elif p['상호'] != p['상호'].upper():
            rep.warn('PARTY-02', f'{역할} 상호에 소문자가 섞여 있다: "{p["상호"]}"',
                     '원산지증명서는 대소문자까지 송장과 일치해야 한다.')
    if not (d.get('수출자', {}).get('통관고유부호')):
        rep.err('PARTY-03', '수출자 통관고유부호가 없다.', '유니패스에서 발급. 없으면 수출신고 불가.')
    if not (d.get('수입자', {}).get('해외거래처부호')):
        rep.err('PARTY-04', '수입자 해외거래처부호가 없다.',
                '미등록이면 수출신고서 작성 자체가 불가하다. 제3자 등록분도 사용 가능하니 먼저 조회할 것.')

    # ── 4. 통관 (08장 / 13장)
    if not (통관.get('물품소재지') or '').strip():
        rep.err('CUS-01', '물품소재지가 없다.',
                '실제 물품이 있는 장소 기준이다. 사업장 주소로 신고하면 조사 대상이 된다.')
    if not (통관.get('적재예정보세구역') or '').strip():
        rep.warn('CUS-02', '적재예정보세구역이 없다.', '관세사에게 반드시 전달해야 하는 두 값 중 하나다.')
    제조자 = (통관.get('제조자') or '').strip()
    if not 제조자 or '미상' in 제조자:
        rep.warn('CUS-03', '제조자가 미상으로 처리되어 있다.',
                 'FTA 원산지증명서 발급이 어려워지고 관세환급이 제한된다.')

    # ── 5. 적재기한 (관세법 제251조)
    수리일 = 통관.get('수출신고수리일')
    if 수리일:
        try:
            base = dt.date.fromisoformat(수리일)
            기한, 밀림 = 적재기한(base)
            남은 = (기한 - today).days
            휴일메모 = f' (30일째 {base + dt.timedelta(days=적재기한일수)} 이 휴일이라 {밀림}일 순연)' if 밀림 else ''
            if 기한.year not in 공휴일_수록연도:
                rep.warn('LOAD-05', f'{기한.year}년 공휴일이 data/공휴일.csv 에 없다. 계산값이 부정확할 수 있다.',
                         '수출신고필증에 인쇄된 적재의무기한이 정본이다. 공휴일표를 갱신할 것.')
            if 남은 < 0:
                rep.err('LOAD-01', f'적재기한이 {-남은}일 지났다. (기한 {기한}){휴일메모}',
                        '200만원 이하 과태료 + 수출신고 수리 취소 가능. 연장은 기한 내에만 신청할 수 있다.')
            elif 남은 <= 7:
                rep.err('LOAD-02', f'적재기한까지 {남은}일 남았다. (기한 {기한}){휴일메모}',
                        '선적이 어려우면 지금 적재기간 연장승인을 신청할 것. 수리일부터 1년 범위에서 연장된다.')
            else:
                rep.info('LOAD-03', f'적재기한 {기한} (D-{남은}){휴일메모}')
        except ValueError:
            rep.err('LOAD-04', f'수출신고수리일 형식이 잘못됐다: {수리일}', 'YYYY-MM-DD')

    # ── 6. 금액
    통화 = 거래.get('통화')
    총액 = sum((it.get('수량') or 0) * (it.get('단가') or 0) for it in 품목)
    if 총액 <= 0:
        rep.err('AMT-01', '총액이 0 이하다.')
    else:
        rep.info('AMT-02', f'총액 {통화} {총액:,.2f}')

    # ── 7. 원산지증명 (09장)
    협정 = (co.get('적용협정') or '')
    if 협정:
        키 = 협정.replace('한-', '').replace(' FTA', '').replace('CEPA', '').strip()
        형식 = co.get('발급형식')
        기관만 = any(k in 협정 for k in 발급형식['기관발급만'])
        자율만 = any(k in 협정 for k in 발급형식['자율발급만'])
        if 기관만 and 형식 == '자율발급':
            rep.err('CO-01', f'{협정} 은 기관발급 협정이다. 자율발급 불가.',
                    '세관(유니패스) 또는 상공회의소에서 발급 신청할 것.')
        if 자율만 and 형식 == '기관발급':
            rep.err('CO-02', f'{협정} 은 자율발급 협정이다.',
                    '수출자가 직접 작성/서명한다. 판정 근거서류는 스스로 보관해야 한다.')
        for k, (pct, base) in 미소기준.items():
            if k in 협정:
                rep.info('CO-03', f'{협정} 미소기준: {pct}% ({base} 기준). 섬유류(50~63류)는 별도.')
                break
        else:
            rep.warn('CO-04', f'{협정} 의 미소기준은 확인되지 않았다.',
                     '협정문 또는 FTA 포털에서 직접 확인할 것.')
        if any(k in 협정 for k in ('EU', '영국')) and not co.get('인증수출자번호'):
            if 통화 == 'EUR' and 총액 > 인증수출자_기준_EUR:
                rep.err('CO-05', f'{협정} 이고 총액 EUR {총액:,.2f} 로 6천유로를 초과한다.',
                        '인증수출자만 원산지신고서를 작성할 수 있다.')
            else:
                rep.warn('CO-06', f'{협정} 은 6천유로 초과 시 인증수출자가 필요하다.',
                         '단일 운송서류 기준 단일 수출자→단일 수하인 총가격으로 판단한다.')
        rows = load_psr()
        if rows:
            for it in 품목:
                r = find_psr(rows, it.get('hs_code'), 키)
                if r:
                    rep.info('CO-07', f'[{it.get("hs_code")}] {키} 원산지결정기준: {r["기준"][:70]}')

    # ── 8. 결제 (03장 / 12장)
    결제조건 = (거래.get('결제조건') or '')
    코드 = 거래.get('결제방법코드')
    if 코드 == 'TT' and not re.search(r'T/T|TT|송금|advance|remittance', 결제조건, re.I):
        rep.warn('PAY-01', f'결제방법코드 TT 인데 결제조건 문구와 맞지 않는다: "{결제조건}"')
    if 코드 in ('LS', 'LU') and not re.search(r'L/C|LC|credit', 결제조건, re.I):
        rep.warn('PAY-02', f'결제방법코드 {코드}(신용장) 인데 결제조건에 L/C 표기가 없다.')
    if re.search(r'\bafter\b|사후', 결제조건, re.I) and not re.search(r'advance|선급|선수금', 결제조건, re.I):
        rep.warn('PAY-03', '전액 사후송금 조건으로 보인다.',
                 '신용위험이 전액 남는다. K-SURE 신용조사와 무역보험을 검토할 것. (14장)')
    if 코드 in ('LS', 'LU'):
        rep.info('PAY-04', '신용장 거래는 Master B/L 이 요구된다. 포워더에게 미리 요청할 것. (10장)')

    return rep


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('file')
    ap.add_argument('--today', default=None, help='기준일 (YYYY-MM-DD). 적재기한 계산용')
    ap.add_argument('--json', action='store_true', help='결과를 JSON 으로 출력')
    a = ap.parse_args()

    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    with open(a.file, encoding='utf-8') as f:
        d = json.load(f)
    rep = validate(d, today)

    if a.json:
        print(json.dumps(rep.rows, ensure_ascii=False, indent=1))
    else:
        mark = {'ERROR': '✖', 'WARN': '!', 'INFO': '·'}
        for lv in ('ERROR', 'WARN', 'INFO'):
            for r in [x for x in rep.rows if x['level'] == lv]:
                print(f"{mark[lv]} [{r['code']}] {r['msg']}")
                if r['fix']:
                    print(f"    → {r['fix']}")
        n_e = len(rep.errors)
        n_w = sum(1 for r in rep.rows if r['level'] == 'WARN')
        print(f"\n오류 {n_e} / 경고 {n_w}")
    return 1 if rep.errors else 0


if __name__ == '__main__':
    sys.exit(main())
