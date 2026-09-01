#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""manual/*.md → docs/수출설명서.html (읽기용 한 페이지).

    python3 scripts/build_manual_html.py

마크다운 변환기는 이 저장소가 실제로 쓰는 문법만 처리한다.
제목 / 표 / 목록 / 인용 / 코드펜스 / 수평선 / 굵게 / 인라인코드 / 링크.
문법을 새로 쓰면 여기에 추가할 것.
"""
import html as H
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MANUAL = os.path.join(ROOT, 'manual')
OUT = os.path.join(ROOT, 'docs', '수출설명서.html')
GITHUB = 'https://github.com/jslee2011070102-ship-it/export-study/blob/claude/export-channel-analysis-4t6j34/'


# ── 인라인 ────────────────────────────────────────────────────────
def 링크변환(대상):
    """마크다운 링크 대상을 페이지 내 앵커나 GitHub URL 로 바꾼다."""
    if 대상.startswith(('http://', 'https://', '#')):
        return 대상, 대상.startswith('http')
    m = re.match(r'^(\d{2})_[^/]+\.md$', 대상)
    if m:
        return f'#ch-{m.group(1)}', False
    m = re.match(r'^부록([A-Z])_[^/]+\.md$', 대상)
    if m:
        return f'#ap-{m.group(1)}', False
    return GITHUB + 대상.replace('../', ''), True


def 인라인(t):
    t = H.escape(t, quote=False)
    t = re.sub(r'`([^`]+)`', lambda m: f'<code>{m.group(1)}</code>', t)
    t = re.sub(r'\*\*([^*]+)\*\*', lambda m: f'<strong>{m.group(1)}</strong>', t)

    def _a(m):
        글, 대상 = m.group(1), m.group(2)
        url, 외부 = 링크변환(대상)
        추가 = ' target="_blank" rel="noopener" class="ext"' if 외부 else ''
        return f'<a href="{H.escape(url, quote=True)}"{추가}>{글}</a>'
    return re.sub(r'\[([^\]]+)\]\(([^)]+)\)', _a, t)


def 슬러그(t):
    t = re.sub(r'<[^>]+>', '', t)
    t = re.sub(r'[^\w가-힣]+', '-', t).strip('-')
    return t[:60] or 'x'


# ── 블록 ──────────────────────────────────────────────────────────
def 변환(md, 문서id):
    줄들 = md.split('\n')
    out, 소제목 = [], []
    i, n = 0, len(줄들)
    쓴앵커 = set()

    def 앵커(t):
        s = f'{문서id}--{슬러그(t)}'
        k, c = s, 2
        while k in 쓴앵커:
            k, c = f'{s}-{c}', c + 1
        쓴앵커.add(k)
        return k

    while i < n:
        L = 줄들[i]

        if not L.strip():
            i += 1
            continue

        if L.startswith('```'):                                   # 코드펜스
            i += 1
            버퍼 = []
            while i < n and not 줄들[i].startswith('```'):
                버퍼.append(줄들[i]); i += 1
            i += 1
            out.append('<pre><code>' + H.escape('\n'.join(버퍼)) + '</code></pre>')
            continue

        m = re.match(r'^(#{1,6})\s+(.*)$', L)                     # 제목
        if m:
            lv, 글 = len(m.group(1)), m.group(2).strip()
            if lv == 1:                                           # 문서 제목은 바깥에서 처리
                i += 1
                continue
            a = 앵커(글)
            if lv == 2:
                소제목.append((a, 글))
            out.append(f'<h{lv} id="{a}">{인라인(글)}'
                       f'<a class="anchor" href="#{a}" aria-label="이 절 링크">#</a></h{lv}>')
            i += 1
            continue

        if re.match(r'^-{3,}\s*$', L):                            # 수평선
            out.append('<hr>')
            i += 1
            continue

        if L.startswith('|'):                                     # 표
            버퍼 = []
            while i < n and 줄들[i].startswith('|'):
                버퍼.append(줄들[i]); i += 1
            out.append(표(버퍼))
            continue

        if L.startswith('> '):                                    # 인용 (법령/원문)
            버퍼 = []
            while i < n and 줄들[i].startswith('>'):
                버퍼.append(줄들[i].lstrip('>').strip()); i += 1
            글 = ' '.join(버퍼)
            출처 = ''
            m2 = re.search(r'\s—\s([^—]+)$', 글)
            if m2:
                글, 출처 = 글[:m2.start()], m2.group(1).strip()
            out.append('<blockquote><p>' + 인라인(글) + '</p>'
                       + (f'<cite>{인라인(출처)}</cite>' if 출처 else '') + '</blockquote>')
            continue

        if re.match(r'^[-*]\s+', L):                              # 목록
            항목 = []
            while i < n and re.match(r'^[-*]\s+', 줄들[i]):
                항목.append(re.sub(r'^[-*]\s+', '', 줄들[i])); i += 1
            out.append('<ul>' + ''.join(f'<li>{인라인(x)}</li>' for x in 항목) + '</ul>')
            continue

        문단 = []                                                  # 문단
        while i < n and 줄들[i].strip() and not re.match(
                r'^(#{1,6}\s|\||>|[-*]\s|```|-{3,}\s*$)', 줄들[i]):
            문단.append(줄들[i].strip()); i += 1
        if 문단:
            out.append('<p>' + 인라인(' '.join(문단)) + '</p>')

    return '\n'.join(out), 소제목


def 표(줄들):
    행 = [[c.strip() for c in r.strip().strip('|').split('|')] for r in 줄들]
    if len(행) < 2 or not all(re.match(r'^:?-{2,}:?$', c) for c in 행[1]):
        머리, 몸 = None, 행
    else:
        머리, 몸 = 행[0], 행[2:]
    폭 = max(len(r) for r in 행)
    숫자열 = set()
    for c in range(폭):                     # 숫자가 많은 열은 tabular-nums + 우측정렬
        값 = [r[c] for r in 몸 if c < len(r) and r[c]]
        if 값 and sum(bool(re.fullmatch(r'[\d,.\s%~\-–]+', re.sub(r'<[^>]+>|\*\*|`', '', v)))
                     for v in 값) >= len(값) * 0.7:
            숫자열.add(c)
    s = ['<div class="tw"><table>']
    if 머리:
        s.append('<thead><tr>' + ''.join(
            f'<th{" class=num" if k in 숫자열 else ""}>{인라인(c)}</th>'
            for k, c in enumerate(머리)) + '</tr></thead>')
    s.append('<tbody>')
    for r in 몸:
        s.append('<tr>' + ''.join(
            f'<td{" class=num" if k in 숫자열 else ""}>{인라인(c)}</td>'
            for k, c in enumerate(r)) + '</tr>')
    s.append('</tbody></table></div>')
    return ''.join(s)


# ── 문서 수집 ──────────────────────────────────────────────────────
def 문서들():
    목록 = []
    for f in sorted(os.listdir(MANUAL)):
        if not f.endswith('.md') or f == 'README.md':
            continue
        m = re.match(r'^(\d{2})_(.+)\.md$', f)
        a = re.match(r'^부록([A-Z])_(.+)\.md$', f)
        if m:
            목록.append(('ch', m.group(1), f))
        elif a:
            목록.append(('ap', a.group(1), f))
    목록.sort(key=lambda x: (x[0] != 'ch', x[1]))
    return 목록


def 만들기():
    본문, 목차 = [], []
    for 종류, 번호, 파일 in 문서들():
        md = open(os.path.join(MANUAL, 파일), encoding='utf-8').read()
        제목 = '제목 없음'
        m = re.search(r'^#\s+(.*)$', md, re.M)
        if m:
            제목 = m.group(1).strip()
        # "00. 수출 1건의 전체 흐름" / "부록 A. 수출신고필증 읽는 법" 에서 이름만
        이름 = re.sub(r'^(?:\d{2}\.|부록\s*[A-Z]\.)\s*', '', 제목).strip()

        몸, 소제목 = 변환(md, f'{종류}-{번호}')
        # 첫 문단을 표제문으로 끌어올린다
        표제 = ''
        m2 = re.search(r'<p>(.*?)</p>', 몸, re.S)
        if m2:
            표제 = m2.group(1)
            몸 = 몸[:m2.start()] + 몸[m2.end():]

        아이디 = f'{종류}-{번호}'
        라벨 = 번호 if 종류 == 'ch' else f'부록 {번호}'
        본문.append(f'''<section class="doc" id="{아이디}">
  <header class="doc-head">
    <div class="num" aria-hidden="true">{라벨}</div>
    <h1>{H.escape(이름)}</h1>
    {f'<p class="stand">{표제}</p>' if 표제 else ''}
  </header>
  {몸}
</section>''')
        목차.append((아이디, 라벨, 이름, 소제목, 종류))
    return 본문, 목차


def 사이드바(목차):
    s = []
    앞종류 = None
    for 아이디, 라벨, 이름, 소제목, 종류 in 목차:
        if 종류 != 앞종류:
            s.append(f'<div class="rail-sec">{"본문" if 종류 == "ch" else "부록"}</div>')
            앞종류 = 종류
        하위 = ''.join(f'<a class="sub" href="#{a}" data-t="{H.escape(t, True)}">{H.escape(t)}</a>'
                      for a, t in 소제목)
        s.append(f'''<div class="rail-item" data-t="{H.escape(라벨 + " " + 이름, True)}">
  <a class="top" href="#{아이디}"><span class="n">{H.escape(라벨)}</span>{H.escape(이름)}</a>
  <div class="subs">{하위}</div>
</div>''')
    return '\n'.join(s)


CSS = '''
:root{
  --paper:#F7F7F4; --card:#FFFFFF; --sunk:#F0F0EC;
  --ink:#191917; --ink-2:#4E4E48; --ink-3:#82827A;
  --rule:#DFDFD8; --rule-soft:#EBEBE5;
  --seal:#A63328; --seal-soft:#F6E9E7;
  --slate:#2C4A5E; --slate-soft:#E7EDF1;
  --shadow:0 1px 2px rgba(25,25,23,.05);
}
@media (prefers-color-scheme:dark){
  :root:not([data-theme="light"]){
    --paper:#141416; --card:#1B1B1E; --sunk:#202024;
    --ink:#E9E9E4; --ink-2:#B4B4AC; --ink-3:#85857D;
    --rule:#313136; --rule-soft:#26262A;
    --seal:#E0796C; --seal-soft:#33201D;
    --slate:#8FB4CA; --slate-soft:#1B2830;
    --shadow:0 1px 2px rgba(0,0,0,.5);
  }
}
:root[data-theme="dark"]{
  --paper:#141416; --card:#1B1B1E; --sunk:#202024;
  --ink:#E9E9E4; --ink-2:#B4B4AC; --ink-3:#85857D;
  --rule:#313136; --rule-soft:#26262A;
  --seal:#E0796C; --seal-soft:#33201D;
  --slate:#8FB4CA; --slate-soft:#1B2830;
  --shadow:0 1px 2px rgba(0,0,0,.5);
}
*{box-sizing:border-box}
html{scroll-behavior:smooth; scroll-padding-top:20px}
@media (prefers-reduced-motion:reduce){ html{scroll-behavior:auto} }
body{
  margin:0; background:var(--paper); color:var(--ink);
  font-family:"IBM Plex Sans KR",system-ui,-apple-system,"Malgun Gothic",sans-serif;
  font-size:15px; line-height:1.78; letter-spacing:-.003em;
  -webkit-font-smoothing:antialiased;
}
.mono,code,.num-f{font-family:"IBM Plex Mono","IBM Plex Sans KR",monospace}

/* ── 레이아웃 ── */
.shell{display:grid; grid-template-columns:284px minmax(0,1fr); gap:0; min-height:100vh}
.rail{
  border-right:1px solid var(--rule); background:var(--card);
  position:sticky; top:0; height:100vh; overflow-y:auto; padding:26px 0 60px;
}
.main{padding:0 40px 140px; min-width:0}
.wrap{max-width:70ch; margin:0 auto}

/* ── 표지 ── */
.cover{padding:70px 0 30px; border-bottom:2px solid var(--ink)}
.cover .kicker{
  font-size:12.5px; font-weight:500; letter-spacing:.02em; color:var(--seal);
}
.cover h1{
  font-family:"Gowun Batang",Georgia,serif; font-weight:700;
  font-size:clamp(32px,5vw,50px); line-height:1.18; letter-spacing:-.02em; margin:14px 0 0;
  text-wrap:balance;
}
.cover p{color:var(--ink-2); font-size:16px; margin:16px 0 0; max-width:62ch; font-weight:300}
.facts{display:flex; flex-wrap:wrap; gap:6px 26px; margin-top:26px; font-size:12.5px; color:var(--ink-3)}
.facts b{color:var(--ink-2); font-weight:500}

/* ── 사이드바 ── */
.rail-head{padding:0 20px 16px; border-bottom:1px solid var(--rule); margin-bottom:14px}
.rail-head .t{font-family:"Gowun Batang",serif; font-weight:700; font-size:16px}
.rail-head .s{font-size:11.5px; color:var(--ink-3); margin-top:2px}
.filter{
  width:100%; margin-top:12px; padding:7px 10px; font:inherit; font-size:13px;
  border:1px solid var(--rule); border-radius:3px; background:var(--paper); color:var(--ink);
}
.filter:focus{outline:2px solid var(--slate); outline-offset:1px; border-color:transparent}
.rail-sec{
  font-size:11.5px; font-weight:500; letter-spacing:.02em;
  color:var(--ink-3); padding:16px 20px 6px;
}
.rail-item.hide{display:none}
a.top{
  display:flex; gap:10px; align-items:baseline; padding:6px 20px; text-decoration:none;
  color:var(--ink-2); font-size:13.5px; border-left:2px solid transparent;
}
a.top .n{
  font-family:"IBM Plex Mono",monospace; font-size:11.5px; color:var(--seal);
  min-width:38px; flex:none; font-variant-numeric:tabular-nums;
}
a.top:hover{background:var(--sunk); color:var(--ink)}
a.top.on{border-left-color:var(--seal); background:var(--seal-soft); color:var(--ink); font-weight:500}
.subs{display:none; padding:2px 0 8px}
.rail-item.open .subs{display:block}
a.sub{
  display:block; padding:3px 20px 3px 68px; font-size:12.5px; color:var(--ink-3);
  text-decoration:none; border-left:2px solid transparent;
}
a.sub:hover{color:var(--ink); background:var(--sunk)}
a.sub.on{color:var(--seal); border-left-color:var(--rule)}

/* ── 문서 ── */
.doc{padding-top:64px; scroll-margin-top:16px}
.doc + .doc{border-top:1px solid var(--rule); margin-top:64px}
.doc-head{margin-bottom:8px}
.doc-head .num{
  font-family:"IBM Plex Mono","IBM Plex Sans KR",monospace; font-size:12px; letter-spacing:.06em;
  color:var(--seal); font-variant-numeric:tabular-nums;
}
.doc-head h1{
  font-family:"Gowun Batang",Georgia,serif; font-weight:700; font-size:30px;
  line-height:1.25; letter-spacing:-.02em; margin:6px 0 0; text-wrap:balance;
}
.stand{
  font-size:15.5px; color:var(--ink-2); margin:14px 0 0; padding-left:14px;
  border-left:3px solid var(--seal); font-weight:300;
}
h2{
  font-family:"Gowun Batang",Georgia,serif; font-size:21px; font-weight:700;
  letter-spacing:-.015em; margin:46px 0 0; padding-bottom:9px;
  border-bottom:1.5px solid var(--ink); text-wrap:balance;
}
h3{font-size:15.5px; font-weight:600; margin:32px 0 0; letter-spacing:-.01em}
h4{font-size:14px; font-weight:600; margin:24px 0 0; color:var(--ink-2)}
p{margin:14px 0}
.main a{color:var(--slate); text-decoration:none; border-bottom:1px solid color-mix(in srgb,var(--slate) 32%,transparent)}
.main a:hover{border-bottom-color:var(--slate)}
a:focus-visible{outline:2px solid var(--slate); outline-offset:2px; border-radius:2px}
.rail a{border-bottom:0}
.main a.ext::after{content:"↗"; font-size:.78em; margin-left:2px; color:var(--ink-3)}
.anchor{
  margin-left:8px; color:var(--ink-3); border:0; font-weight:400; font-size:.7em;
  opacity:0; transition:opacity .12s;
}
h2:hover .anchor,h3:hover .anchor{opacity:1}
strong{font-weight:600; color:var(--ink)}
code{
  font-size:.87em; background:var(--sunk); padding:1px 5px; border-radius:3px;
  border:1px solid var(--rule-soft);
}
ul{margin:14px 0; padding-left:20px}
li{margin:6px 0; color:var(--ink-2)}
li strong{color:var(--ink)}
hr{border:0; border-top:1px solid var(--rule); margin:36px 0}
pre{
  background:var(--sunk); border:1px solid var(--rule); border-radius:3px;
  padding:16px 18px; overflow-x:auto; margin:20px 0; line-height:1.5;
}
pre code{background:none; border:0; padding:0; font-size:12.5px}
blockquote{
  margin:20px 0; padding:14px 18px; background:var(--slate-soft);
  border-left:3px solid var(--slate); border-radius:0 3px 3px 0;
}
blockquote p{margin:0; font-size:14.5px; color:var(--ink-2)}
blockquote cite{
  display:block; margin-top:8px; font-style:normal;
  font-family:"IBM Plex Mono",monospace; font-size:11.5px; color:var(--slate);
}

/* ── 표: 서식의 격자 ── */
.tw{overflow-x:auto; margin:20px 0; border:1px solid var(--rule); border-radius:3px; background:var(--card)}
table{width:100%; border-collapse:collapse; font-size:13.5px}
th{
  text-align:left; font-weight:500; font-size:11.5px; letter-spacing:.02em;
  color:var(--ink-3); padding:10px 14px; background:var(--sunk);
  border-bottom:1px solid var(--rule); white-space:nowrap;
}
td{padding:10px 14px; border-bottom:1px solid var(--rule-soft); vertical-align:top; color:var(--ink-2)}
td strong{color:var(--ink)}
tr:last-child td{border-bottom:0}
th.num,td.num{font-variant-numeric:tabular-nums; white-space:nowrap}
td.num{font-family:"IBM Plex Mono","IBM Plex Sans KR",monospace; font-size:12.5px}

/* ── 상단 바 (모바일) ── */
.bar{display:none}
@media (max-width:900px){
  .shell{grid-template-columns:1fr}
  .rail{
    position:fixed; inset:0 auto 0 0; width:290px; z-index:40; transform:translateX(-100%);
    transition:transform .2s ease; box-shadow:var(--shadow);
  }
  .rail.open{transform:none}
  .main{padding:0 20px 120px}
  .bar{
    display:flex; gap:12px; align-items:center; position:sticky; top:0; z-index:30;
    background:color-mix(in srgb,var(--paper) 92%,transparent); backdrop-filter:blur(8px);
    border-bottom:1px solid var(--rule); padding:10px 0; margin:0 -20px; padding-left:20px;
  }
  .bar button{
    font:inherit; font-size:13px; padding:6px 12px; border:1px solid var(--rule);
    background:var(--card); color:var(--ink); border-radius:3px; cursor:pointer;
  }
  .bar .cur{font-size:12.5px; color:var(--ink-3); overflow:hidden; text-overflow:ellipsis; white-space:nowrap}
  .cover{padding-top:34px}
}
@media print{
  .rail,.bar,.anchor{display:none}
  .shell{display:block}
  .main{padding:0}
  .doc{page-break-before:always; padding-top:0}
  body{font-size:10.5pt; background:#fff; color:#000}
}
'''

JS = '''
(function(){
  var rail=document.querySelector('.rail'), f=document.getElementById('filter');
  var tops=[].slice.call(document.querySelectorAll('a.top'));
  var subs=[].slice.call(document.querySelectorAll('a.sub'));
  var cur=document.getElementById('cur');

  f.addEventListener('input', function(){
    var q=f.value.trim().toLowerCase();
    document.querySelectorAll('.rail-item').forEach(function(it){
      var hit = !q || it.dataset.t.toLowerCase().indexOf(q)>-1 ||
        [].some.call(it.querySelectorAll('a.sub'), function(s){
          return s.dataset.t.toLowerCase().indexOf(q)>-1; });
      it.classList.toggle('hide', !hit);
      if(q) it.classList.toggle('open', hit);
    });
  });

  function 표시(id){
    tops.forEach(function(a){
      var on = a.getAttribute('href')==='#'+id;
      a.classList.toggle('on', on);
      a.closest('.rail-item').classList.toggle('open', on);
      if(on){
        if(cur){ var nEl=a.querySelector('.n');
          cur.textContent = (nEl? nEl.textContent.trim()+'  ' : '') +
            a.textContent.replace(nEl? nEl.textContent : '', '').trim(); }
        var r=a.getBoundingClientRect(), rr=rail.getBoundingClientRect();
        if(r.top<rr.top+40 || r.bottom>rr.bottom-40) a.scrollIntoView({block:'center'});
      }
    });
  }
  var io=new IntersectionObserver(function(es){
    es.forEach(function(e){ if(e.isIntersecting) 표시(e.target.id); });
  },{rootMargin:'-10% 0px -80% 0px', threshold:0});
  document.querySelectorAll('section.doc').forEach(function(s){ io.observe(s); });

  var io2=new IntersectionObserver(function(es){
    es.forEach(function(e){
      if(!e.isIntersecting) return;
      subs.forEach(function(a){ a.classList.toggle('on', a.getAttribute('href')==='#'+e.target.id); });
    });
  },{rootMargin:'-12% 0px -78% 0px', threshold:0});
  document.querySelectorAll('h2[id]').forEach(function(h){ io2.observe(h); });

  var t=document.getElementById('toggle');
  if(t) t.addEventListener('click', function(){ rail.classList.toggle('open'); });
  rail.addEventListener('click', function(e){
    if(e.target.closest('a') && window.innerWidth<=900) rail.classList.remove('open');
  });
})();
'''


def main():
    본문, 목차 = 만들기()
    글자수 = sum(len(open(os.path.join(MANUAL, f), encoding='utf-8').read())
                for _, _, f in 문서들())
    장수 = sum(1 for x in 목차 if x[4] == 'ch')
    부록수 = len(목차) - 장수

    페이지 = f'''<title>수출설명서</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Gowun+Batang:wght@400;700&family=IBM+Plex+Sans+KR:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>{CSS}</style>

<div class="shell">
<nav class="rail" id="rail">
  <div class="rail-head">
    <div class="t">수출설명서</div>
    <div class="s">{장수}개 장 · 부록 {부록수}종</div>
    <input class="filter" id="filter" type="search" placeholder="장·절 이름으로 찾기" aria-label="목차 검색">
  </div>
  {사이드바(목차)}
</nav>

<div class="main">
  <div class="bar">
    <button id="toggle" aria-label="목차 열기">목차</button>
    <span class="cur" id="cur"></span>
  </div>
  <div class="wrap">
    <header class="cover">
      <div class="kicker">첫 수출 1건을 스스로 끝내기 위한 실무 설명서</div>
      <h1>수출설명서</h1>
      <p>무역전문채널 <b>무꿈사TV</b> 전수 분석에서 출발해, 법령 원문으로 수치와 기한을 대조하고
      실무 순서대로 다시 쓴 것. 각 장은 <b>이 장이 끝나면 무엇이 손에 남는지</b>로 시작하고,
      수출자가 직접 할 일과 관세사·포워더에게 넘길 일을 구분한 뒤 체크리스트로 끝난다.</p>
      <div class="facts">
        <span>본문 <b>{장수}개 장</b></span>
        <span>부록 <b>{부록수}종</b></span>
        <span>분량 <b>{글자수:,}자</b></span>
        <span>법령 검증 <b>2026-08-31</b></span>
      </div>
    </header>
    {chr(10).join(본문)}
  </div>
</div>
</div>
<script>{JS}</script>
'''
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(페이지)
    print(f'생성: {OUT}')
    print(f'  본문 {장수}개 장 + 부록 {부록수}종 / 원문 {글자수:,}자 / HTML {len(페이지):,}자')


if __name__ == '__main__':
    main()
