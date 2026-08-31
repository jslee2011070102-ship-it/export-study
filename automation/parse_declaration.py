#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""수출신고필증(면장) 파서.

유니패스에서 발급된 전자 수출신고필증 PDF 에서 항목을 뽑아 JSON 으로 만든다.
항목 번호/이름은 「수출 및 반송통관에 관한 고시」 별지 제1호서식 및
별표 제1호 수출신고서 작성요령 기준.

    python3 automation/parse_declaration.py 필증.pdf
    python3 automation/parse_declaration.py 필증.pdf -o 필증.json
    python3 automation/parse_declaration.py 필증.pdf --check automation/samples/sample.json

PDF 는 pdftotext -layout (poppler-utils) 으로 텍스트를 뽑는다. 이미 뽑아둔 .txt 도 받는다.
스캔본(이미지 PDF)은 대상이 아니다. 유니패스 발급본은 텍스트 레이어가 있다.
"""
import argparse, csv, datetime as dt, json, os, re, shutil, subprocess, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
공휴일표 = os.path.join(ROOT, 'data', '공휴일.csv')

적재기한일수 = 30          # 관세법 제251조제1항
인코텀즈_11 = ('EXW', 'FCA', 'FAS', 'FOB', 'CPT', 'CIP', 'CFR', 'CIF', 'DAP', 'DPU', 'DDP')

# ── 텍스트 추출 ────────────────────────────────────────────────
def 텍스트(path):
    if path.lower().endswith('.txt'):
        return open(path, encoding='utf-8').read()
    if not shutil.which('pdftotext'):
        sys.exit('pdftotext 가 없다. poppler-utils 를 설치할 것. (apt-get install poppler-utils)')
    r = subprocess.run(['pdftotext', '-layout', path, '-'],
                       capture_output=True, text=True)
    if r.returncode != 0:
        sys.exit(f'pdftotext 실패: {r.stderr.strip()}')
    if len(r.stdout.strip()) < 200:
        sys.exit('텍스트가 거의 없다. 스캔 이미지 PDF 로 보인다. 유니패스 원본을 받을 것.')
    return r.stdout


# ── 항목 추출 ──────────────────────────────────────────────────
def _1(pat, t, flags=0, g=1):
    m = re.search(pat, t, flags)
    return m.group(g).strip() if m else None

def _num(s):
    if s is None or s.strip() == '':
        return None
    s = s.replace(',', '').replace('￦', '').replace('$', '').strip()
    try:
        return float(s) if '.' in s else int(s)
    except ValueError:
        return None

def _date(s):
    if not s:
        return None
    s = s.strip().replace('/', '-')
    try:
        return dt.date.fromisoformat(s).isoformat()
    except ValueError:
        return None


def 헤더(t):
    d = {}
    d['문서종류'] = _1(r'((?:간이\)?)?수출신고필증\s*\([^)]*\)|수출신고서(?:\(을지\))?)', t) or '미상'
    d['전자문서'] = '전자문서' in t

    # 5 신고번호 / 6 세관.과 / 7 신고일자 / 8 신고구분 / 9 C/S구분 - 한 줄에 같이 찍힌다
    m = re.search(r'(\d{5}-\d{2}-\d{7}\w?)\s+(\S+)\s+(\d{4}-\d{2}-\d{2})\s+(\S+)\s+(\S+)\s*$',
                  t, re.M)
    if m:
        d['신고번호'], d['세관과'] = m.group(1), m.group(2)
        d['신고일자'], d['신고구분'], d['CS구분'] = _date(m.group(3)), m.group(4), m.group(5)
    else:
        d['신고번호'] = _1(r'신고번호\s*(\d{5}-\d{2}-\d{7}\w?)', t)
        d['신고일자'] = _date(_1(r'신고일자\s*(\d{4}[-/]\d{2}[-/]\d{2})', t))

    d['신고자']    = _1(r'^\s*1\s*신고자\s+(.+?)\s{2,}', t, re.M)
    d['수출대행자'] = _1(r'수출대행자\s*\n?\s*(.+?)\s*$', t, re.M)
    d['수출화주']  = _1(r'수\s*출\s*화\s*주\s+(.+?)\s{2,}', t)
    d['통관고유부호'] = _1(r'\(통관고유부호\)\s*(\S+)', t)
    d['사업자등록번호'] = _1(r'\(사업자등록번호\)\s*(\d{3}-\d{2}-\d{5})', t)
    d['대표자']   = _1(r'\(대표자\)\s*(\S+)', t)
    d['제조자']   = _1(r'^\s*3\s*제\s*조\s*자\s+(.+?)\s{2,}', t, re.M)
    d['구매자']   = _1(r'^\s*4\s*구\s*매\s*자\s+(.+?)\s{2,}', t, re.M)
    d['구매자부호'] = _1(r'\(구매자부호\)\s*(\S+)', t)

    d['거래구분'] = _1(r'거래구분\s+(\S+)', t)
    d['종류']    = _1(r'종류\s+(\S+)', t)
    d['결제방법'] = _1(r'결제방법\s+(\S+)', t)
    d['목적국']  = _1(r'목적국\s+([A-Z]{2})\b', t)
    d['적재항']  = _1(r'적재항\s+(\S+)', t)
    d['운송형태'] = _1(r'운송형태\s+(\S+)', t)
    d['물품상태'] = _1(r'물품상태\s+(\S+)', t)
    d['환급신청인'] = _1(r'환급신청인\s+(\S+)', t)
    d['적재예정보세구역'] = _1(r'적재예정보세구역\s*\n?\s*.*?(\d{8})', t)
    d['LC번호']  = _1(r'L\s*/\s*C\s*번\s*호\s+(\S+)', t)

    # 44~49
    d['총중량_kg']   = _num(_1(r'총중량\s+([\d,\.]+)\s*\(', t))
    d['총포장갯수']   = _num(_1(r'총포장갯수\s+([\d,]+)\s*\(', t))
    d['포장종류']    = _1(r'총포장갯수\s+[\d,]+\s*\((\w+)\)', t)
    m = re.search(r'총신고가격\s*.*?\$\s*([\d,]+).*?￦\s*([\d,]+)', t, re.S)
    if m:
        d['총신고가격_USD'] = _num(m.group(1))
        d['총신고가격_KRW'] = _num(m.group(2))
    d['운임_KRW']   = _num(_1(r'운임\(￦\)\s*([\d,]*)', t))
    d['보험료_KRW'] = _num(_1(r'보험료\(￦\)\s*([\d,]*)', t))

    # 49 결제금액 = 인도조건-통화종류-금액
    m = re.search(r'결제금액\s+([A-Z]{3})-([A-Z]{3})-([\d,\.]+)', t)
    if m:
        d['인도조건'], d['통화'] = m.group(1), m.group(2)
        d['결제금액'] = _num(m.group(3))

    d['컨테이너번호'] = _1(r'컨테이너번호\s+([A-Z]{4}\d{7})', t)
    d['적재의무기한'] = _date(_1(r'적재의무기한\s+(\d{4}/\d{2}/\d{2})', t))
    d['신고수리일자'] = _date(_1(r'신고수리일자\s+(\d{4}/\d{2}/\d{2})', t))
    d['발행번호']   = _1(r'발\s*행\s*번\s*호\s*:\s*(\d+)', t)
    return {k: v for k, v in d.items() if v not in (None, '')}


def 란(t):
    """품명·규격 블록(란)을 뽑는다. 갑지/을지에 걸쳐 같은 란번호가 반복되므로 병합한다."""
    블록 = re.split(r'●?\s*품명\s*[․·ㆍ]\s*규격\s*\(란번호/총란수\s*:', t)[1:]
    출력 = {}
    for b in 블록:
        m = re.match(r'\s*(\d+)\s*/\s*(\d+)\s*\)', b)
        if not m:
            continue
        no, 총란수 = int(m.group(1)), int(m.group(2))
        r = 출력.setdefault(no, {'란번호': no, '총란수': 총란수, '모델규격': []})
        for k, pat in [
            ('품명',     r'품\s*명\s+(.+?)\s*$'),
            ('거래품명',  r'거래품명\s+(.+?)\s*$'),
            ('상표명',    r'상표명\s+(.+?)\s*$'),
            ('세번부호',  r'세번부호\s+(\d{4}\.\d{2}-\d{4})'),
            ('송품장부호', r'송품장부호\s+(\S+)'),
            ('원산지',    r'원산지\s+(\S+)'),
        ]:
            v = _1(pat, b, re.M)
            if v and k not in r:
                r[k] = v
        for k, pat in [('순중량_kg', r'순중량\s+([\d,\.]+)\s*\('),
                       ('포장갯수',  r'포장갯수\(종류\)\s+([\d,]+)\s*\(')]:
            v = _num(_1(pat, b))
            if v is not None and k not in r:
                r[k] = v
        m2 = re.search(r'신고가격\(FOB\)\s*.*?\$\s*([\d,]+).*?￦\s*([\d,]+)', b, re.S)
        if m2 and '신고가격_KRW' not in r:
            r['신고가격_USD'] = _num(m2.group(1))
            r['신고가격_KRW'] = _num(m2.group(2))

        # 모델·규격 행: (NO.nn) ... 수량 (단위)  단가  금액  / 다음 줄에 규격 문자열
        lines = b.splitlines()
        for i, line in enumerate(lines):
            mm = re.match(r'\s*\(NO\.(\d+)\)\s+.*?([\d,]+)\s*\((\w+)\)\s+([\d,\.]+)\s+([\d,\.]+)\s*$', line)
            if not mm:
                continue
            규격 = ''
            for nxt in lines[i + 1:i + 3]:
                s = nxt.strip()
                if s and not re.match(r'\(NO\.\d+\)|발\s*행\s*번\s*호|\d+란|001란', s):
                    규격 = s
                    break
            r['모델규격'].append({
                '순번': int(mm.group(1)), '규격': 규격,
                '수량': _num(mm.group(2)), '단위': mm.group(3),
                '단가': _num(mm.group(4)), '금액': _num(mm.group(5)),
            })
    for r in 출력.values():
        r['모델규격'].sort(key=lambda x: x['순번'])
    return [출력[k] for k in sorted(출력)]


# ── 적재기한 계산 (관세법 제251조 + 제8조제3항) ───────────────────
def _공휴일():
    s = set()
    미검증 = set()
    if os.path.exists(공휴일표):
        with open(공휴일표, encoding='utf-8') as f:
            for row in csv.DictReader(f):
                s.add(dt.date.fromisoformat(row['일자']))
                if row.get('검증', '').strip() == '검증':
                    미검증.add(dt.date.fromisoformat(row['일자']))
    return s, 미검증

공휴일, 공휴일_미검증 = _공휴일()
공휴일_수록연도 = {d.year for d in 공휴일}


def 적재기한(수리일):
    """수리일부터 30일. 말일이 토/일/공휴일이면 관세법 제8조제3항에 따라 다음 날로 밀린다."""
    d = 수리일 + dt.timedelta(days=적재기한일수)
    while d.weekday() >= 5 or d in 공휴일:
        d += dt.timedelta(days=1)
    return d


# ── 상업송장 마스터와 대조 ───────────────────────────────────────
def 대조(필증, master):
    거래 = master.get('거래', {})
    품목 = master.get('품목', [])
    out = []
    def eq(code, 이름, a, b, fix=''):
        if a is None or b in (None, ''):
            return
        ok = str(a).strip().upper() == str(b).strip().upper()
        out.append({'level': 'OK' if ok else 'ERROR', 'code': code,
                    'msg': f'{이름}: 필증 {a} / 송장 {b}', 'fix': '' if ok else fix})

    조건 = (거래.get('가격조건') or '').split()
    eq('DEC-01', '인도조건', 필증.get('인도조건'), 조건[0] if 조건 else None,
       '수출신고 인도조건과 송장 가격조건이 다르면 신고 정정 대상이다.')
    eq('DEC-02', '통화', 필증.get('통화'), 거래.get('통화'))
    eq('DEC-03', '송품장부호', (필증.get('란') or [{}])[0].get('송품장부호'), 거래.get('invoice_no'))

    총액 = sum((i.get('수량') or 0) * (i.get('단가') or 0) for i in 품목)
    결제 = 필증.get('결제금액')
    if 결제 is not None and 총액:
        차 = abs(결제 - 총액)
        out.append({'level': 'OK' if 차 < 0.01 else 'ERROR', 'code': 'DEC-04',
                    'msg': f'결제금액: 필증 {결제:,.2f} / 송장 합계 {총액:,.2f}',
                    'fix': '' if 차 < 0.01 else '송장 총액과 신고 결제금액이 일치해야 한다.'})

    송장중량 = sum((i.get('총중량_kg') or 0) for i in 품목)
    필증중량 = 필증.get('총중량_kg')
    if 필증중량 and 송장중량:
        차 = abs(필증중량 - 송장중량)
        out.append({'level': 'OK' if 차 < 0.5 else 'WARN', 'code': 'DEC-05',
                    'msg': f'총중량: 필증 {필증중량} / 송장 {송장중량}',
                    'fix': '' if 차 < 0.5 else '중량 불일치는 적하목록 대사 단계에서 미선적 처리의 원인이 된다.'})

    필증세번 = {r.get('세번부호', '').replace('.', '').replace('-', '')[:6]
                for r in 필증.get('란', []) if r.get('세번부호')}
    송장세번 = {(i.get('hs_code') or '').replace('.', '')[:6] for i in 품목 if i.get('hs_code')}
    if 필증세번 and 송장세번:
        누락 = 송장세번 - 필증세번
        out.append({'level': 'OK' if not 누락 else 'WARN', 'code': 'DEC-06',
                    'msg': f'세번(6단위): 필증 {sorted(필증세번)} / 송장 {sorted(송장세번)}',
                    'fix': '' if not 누락 else 'FTA 원산지결정기준은 필증 세번 기준으로 판정된다.'})
    return out


def 적재기한_점검(필증, today):
    out = []
    수리 = 필증.get('신고수리일자')
    기한 = 필증.get('적재의무기한')
    if not 기한:
        return out
    기한d = dt.date.fromisoformat(기한)
    if 수리:
        수리d = dt.date.fromisoformat(수리)
        계산 = 적재기한(수리d)
        if 기한d.year not in 공휴일_수록연도:
            out.append({'level': 'INFO', 'code': 'DEC-10',
                        'msg': f'{기한d.year}년 공휴일이 data/공휴일.csv 에 없다. 계산값은 참고만 할 것.',
                        'fix': '필증에 인쇄된 적재의무기한이 정본이다.'})
        elif 계산 != 기한d:
            out.append({'level': 'INFO', 'code': 'DEC-11',
                        'msg': f'계산 적재기한 {계산} ≠ 필증 표기 {기한d}',
                        'fix': '필증 표기가 정본이다. data/공휴일.csv 갱신이 필요할 수 있다.'})
        else:
            out.append({'level': 'OK', 'code': 'DEC-11',
                        'msg': f'적재기한 계산 일치 ({수리d} +{적재기한일수}일 → {계산})', 'fix': ''})
    남은 = (기한d - today).days
    if 남은 < 0:
        out.append({'level': 'ERROR', 'code': 'DEC-12',
                    'msg': f'적재기한이 {-남은}일 지났다. (기한 {기한d})',
                    'fix': '관세법 제251조 - 수출신고 수리가 취소되고 과태료가 부과될 수 있다.'})
    elif 남은 <= 7:
        out.append({'level': 'ERROR', 'code': 'DEC-12',
                    'msg': f'적재기한까지 {남은}일 남았다. (기한 {기한d})',
                    'fix': '연장은 기한 내에만 신청 가능하다. 별지 제14호서식으로 적재기간 연장승인을 신청할 것.'})
    else:
        out.append({'level': 'INFO', 'code': 'DEC-12',
                    'msg': f'적재기한 {기한d} (D-{남은})', 'fix': ''})
    return out


def 자체점검(필증):
    out = []
    조건 = 필증.get('인도조건')
    if 조건 and 조건 not in 인코텀즈_11:
        out.append({'level': 'ERROR', 'code': 'DEC-20',
                    'msg': f'인도조건 "{조건}" 은 인코텀즈 2020 의 11개 규칙이 아니다.',
                    'fix': '별표1 작성요령 - EXW/FAS/FCA/FOB/CFR/CIF/CPT/CIP/DAP/DPU/DDP 만 사용.'})
    결제, 운임, 보험 = 필증.get('결제금액'), 필증.get('운임_KRW'), 필증.get('보험료_KRW')
    if 조건 in ('CFR', 'CPT', 'CIF', 'CIP') and not 운임:
        out.append({'level': 'ERROR', 'code': 'DEC-21',
                    'msg': f'{조건} 인데 운임란이 비어 있다.',
                    'fix': '작성요령 - 결제금액에 운임이 포함된 경우 운임을 원화로 기재해야 한다.'})
    if 조건 in ('CIF', 'CIP') and not 보험:
        out.append({'level': 'ERROR', 'code': 'DEC-22',
                    'msg': f'{조건} 인데 보험료란이 비어 있다.',
                    'fix': '작성요령 - 결제금액에 보험료가 포함된 경우 보험료를 원화로 기재해야 한다.'})
    if 필증.get('제조자') == '미상' or (필증.get('통관고유부호') or '').startswith('제조미상'):
        out.append({'level': 'WARN', 'code': 'DEC-23', 'msg': '제조자가 미상으로 신고되어 있다.',
                    'fix': 'FTA 원산지증명서 발급이 어려워지고 관세환급이 제한된다.'})
    합 = sum((m.get('금액') or 0) for r in 필증.get('란', []) for m in r.get('모델규격', []))
    if 합 and 필증.get('총신고가격_KRW'):
        차 = abs(합 - 필증['총신고가격_KRW'])
        if 차 > 1:
            out.append({'level': 'WARN', 'code': 'DEC-24',
                        'msg': f'모델·규격 금액 합계 {합:,} ≠ 총신고가격(FOB) {필증["총신고가격_KRW"]:,}',
                        'fix': '인도조건이 FOB 가 아니면 운임/보험료 공제 때문에 정상적으로 다를 수 있다.'})
        else:
            out.append({'level': 'OK', 'code': 'DEC-24',
                        'msg': f'모델·규격 금액 합계 = 총신고가격 {합:,}원', 'fix': ''})
    return out


# ── 실행 ───────────────────────────────────────────────────────
def parse(path):
    t = 텍스트(path)
    d = 헤더(t)
    d['란'] = 란(t)
    return d


def main():
    ap = argparse.ArgumentParser(description='수출신고필증 파서')
    ap.add_argument('필증', help='수출신고필증 PDF 또는 pdftotext -layout 결과 txt')
    ap.add_argument('-o', '--out', help='JSON 저장 경로')
    ap.add_argument('--check', help='대조할 상업송장 마스터 JSON (automation/schema.json 형식)')
    ap.add_argument('--today', help='기준일 YYYY-MM-DD (적재기한 시뮬레이션)')
    ap.add_argument('--json', action='store_true', help='점검 결과를 JSON 으로 출력')
    a = ap.parse_args()

    today = dt.date.fromisoformat(a.today) if a.today else dt.date.today()
    d = parse(a.필증)

    if a.out:
        with open(a.out, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
        print(f'저장: {a.out}')

    결과 = 자체점검(d) + 적재기한_점검(d, today)
    if a.check:
        결과 += 대조(d, json.load(open(a.check, encoding='utf-8')))

    if a.json:
        print(json.dumps({'필증': d, '점검': 결과}, ensure_ascii=False, indent=2))
        return 1 if any(r['level'] == 'ERROR' for r in 결과) else 0

    결제 = ''
    if d.get('결제금액') is not None:
        결제 = f"  {d.get('인도조건')}-{d.get('통화')}-{d.get('결제금액'):,.2f}"
    print(f"\n[{d.get('문서종류')}]  신고번호 {d.get('신고번호')}  "
          f"수리 {d.get('신고수리일자')}{결제}")
    print(f"  수출화주 {d.get('수출화주')} / 제조자 {d.get('제조자')} / 구매자 {d.get('구매자')}")
    가격 = ''
    if d.get('총신고가격_KRW') is not None:
        가격 = f"  총신고가격(FOB) ￦{d['총신고가격_KRW']:,} / ${d.get('총신고가격_USD', 0):,}"
    print(f"  목적국 {d.get('목적국')}  적재항 {d.get('적재항')}{가격}")
    for r in d['란']:
        print(f"  [{r['란번호']}/{r['총란수']}] {r.get('품명')}  세번 {r.get('세번부호')}  "
              f"순중량 {r.get('순중량_kg')}kg  모델·규격 {len(r['모델규격'])}행")
    print()
    아이콘 = {'ERROR': '✖', 'WARN': '!', 'INFO': '·', 'OK': '✓'}
    for r in 결과:
        print(f"{아이콘[r['level']]} [{r['code']}] {r['msg']}")
        if r['fix']:
            print(f"    → {r['fix']}")
    n = sum(1 for r in 결과 if r['level'] == 'ERROR')
    w = sum(1 for r in 결과 if r['level'] == 'WARN')
    print(f'\n오류 {n} / 경고 {w}')
    return 1 if n else 0


if __name__ == '__main__':
    sys.exit(main())
