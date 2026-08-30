#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""무꿈사TV 영상 자막을 옵시디언 클리핑 포맷(.md)으로 일괄 저장.

Obsidian Web Clipper 로 한 편씩 긁는 것과 동일한 결과를 64편 자동으로 만든다.
YouTube 를 정상적으로 볼 수 있는 PC 에서 실행할 것.

설치:
    pip install yt-dlp

사용:
    python fetch_transcripts.py --urls data/need_transcript_urls.txt --out "C:/vault/raw/clippings/무역"
    python fetch_transcripts.py --urls data/need_transcript_urls.txt --out ./out --limit 1   # 먼저 1편만 시험

이미 있는 파일은 건너뛴다. 중간에 끊겨도 다시 돌리면 이어서 받는다.
"""
import argparse, json, os, re, sys, time, urllib.request

BLOCK_SEC = 25          # 자막을 묶는 단위. 옵시디언 클리퍼와 동일한 체감 분량
LANGS = ["ko", "ko-orig", "ko-KR"]
WIN_FORBIDDEN = r'[<>:"/\\|?*\x00-\x1f]'


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def hhmmss(sec):
    """0:07 / 12:28 / 1:02:15 형태. 옵시디언 클리핑과 동일."""
    sec = int(sec)
    h, rem = divmod(sec, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


def escape_md(text):
    """설명란 원문의 마크다운 특수문자를 클리퍼와 같은 방식으로 이스케이프."""
    text = re.sub(r'([\\_\[\]])', r'\\\1', text)
    # 구분선(------)이 setext 제목으로 렌더되지 않도록 선두만 이스케이프
    return re.sub(r'^(\s*)([-=]{3,}\s*)$', r'\1\\\2', text)


def safe_filename(title):
    name = re.sub(WIN_FORBIDDEN, ' ', title)
    name = re.sub(r'\s+', ' ', name).strip().rstrip('.')
    return (name[:120] or 'untitled') + '.md'


def pick_caption_url(info):
    """수동자막 우선, 없으면 자동생성자막. json3 포맷을 고른다."""
    for source in (info.get('subtitles') or {}, info.get('automatic_captions') or {}):
        for lang in LANGS:
            for track in source.get(lang) or []:
                if track.get('ext') == 'json3':
                    return track['url']
    return None


def fetch_events(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    data = json.loads(urllib.request.urlopen(req, timeout=60).read())
    out = []
    for ev in data.get('events') or []:
        segs = ev.get('segs') or []
        text = ''.join(s.get('utf8', '') for s in segs).replace('\n', ' ').strip()
        if text:
            out.append((ev.get('tStartMs', 0) / 1000.0, text))
    return out


def group_events(events, block_sec=BLOCK_SEC):
    """자막 조각을 block_sec 단위 블록으로 묶는다. (시작초, 문장) 리스트 반환."""
    blocks, start, buf = [], None, []
    for t, text in events:
        if start is None:
            start = t
        elif t - start >= block_sec:
            blocks.append((start, ' '.join(buf)))
            start, buf = t, []
        buf.append(text)
    if buf:
        blocks.append((start, ' '.join(buf)))
    return blocks


def parse_chapters(info):
    """yt-dlp 가 뽑은 챕터. 없으면 설명란의 타임스탬프 목록에서 직접 파싱."""
    chapters = [(c['start_time'], c['title']) for c in (info.get('chapters') or [])]
    if chapters:
        return chapters
    for line in (info.get('description') or '').split('\n'):
        m = re.match(r'\s*(?:(\d{1,2}):)?(\d{1,2}):(\d{2})\s+(.+?)\s*$', line)
        if m:
            h, mm, ss = int(m.group(1) or 0), int(m.group(2)), int(m.group(3))
            chapters.append((h * 3600 + mm * 60 + ss, m.group(4).strip()))
    return chapters


def build_note(info, blocks, chapters, created):
    vid = info['video_id']
    watch = f"https://www.youtube.com/watch?v={vid}"
    title = info.get('title') or vid
    upload = info.get('upload_date') or ''
    published = f"{upload[:4]}-{upload[4:6]}-{upload[6:8]}" if len(upload) == 8 else ''
    desc = info.get('description') or ''
    # 프론트매터 description 은 개행 없는 앞부분만 (클리퍼와 동일)
    fm_desc = re.sub(r'\s*\n\s*', '', desc)[:300].replace('"', "'")

    out = ['---',
           f'title: "{title}"',
           f'source: "{watch}"',
           'author:',
           f'  - "[[{info.get("uploader") or "무역전문채널 무꿈사TV"}]]"',
           f'published: {published}',
           f'created: {created}',
           f'description: "{fm_desc}"',
           'tags:',
           '  - "clippings"',
           '---',
           f'![]({watch})',
           '']
    for line in desc.split('\n'):
        out.append(escape_md(line) + '  ')
    out.append('')

    pending = sorted(chapters)
    for start, text in blocks:
        while pending and pending[0][0] <= start:
            out += [f'### {escape_md(pending.pop(0)[1])}', '']
        out += [f'**{hhmmss(start)}** · {text}', '']
    for _, name in pending:                      # 자막보다 뒤에 남은 챕터
        out += [f'### {escape_md(name)}', '']
    return '\n'.join(out).rstrip() + '\n'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--urls', required=True, help='URL 목록 텍스트 파일 (한 줄에 하나)')
    ap.add_argument('--out', required=True, help='저장할 폴더 (옵시디언 볼트 안 경로 권장)')
    ap.add_argument('--limit', type=int, default=0, help='앞에서 N개만 처리 (시험용)')
    ap.add_argument('--sleep', type=float, default=2.0, help='영상 사이 대기 초. 차단 방지')
    ap.add_argument('--block-sec', type=int, default=BLOCK_SEC)
    args = ap.parse_args()

    try:
        from yt_dlp import YoutubeDL
    except ImportError:
        raise SystemExit('yt-dlp 가 없다. 먼저 실행: pip install yt-dlp')

    urls = [l.strip() for l in open(args.urls, encoding='utf-8') if l.strip()
            and not l.startswith('#')]
    if args.limit:
        urls = urls[:args.limit]
    os.makedirs(args.out, exist_ok=True)
    created = time.strftime('%Y-%m-%d')

    ydl = YoutubeDL({'skip_download': True, 'quiet': True, 'no_warnings': True,
                     'writesubtitles': True, 'writeautomaticsub': True,
                     'subtitleslangs': LANGS, 'extract_flat': False})

    ok = skipped = failed = 0
    for i, url in enumerate(urls, 1):
        try:
            info = ydl.extract_info(url, download=False)
        except Exception as exc:                                  # noqa: BLE001
            log(f'[{i}/{len(urls)}] 메타데이터 실패 {url}: {exc}')
            failed += 1
            continue

        path = os.path.join(args.out, safe_filename(info.get('title') or info['id']))
        if os.path.exists(path):
            log(f'[{i}/{len(urls)}] 건너뜀 (이미 있음): {os.path.basename(path)}')
            skipped += 1
            continue

        cap_url = pick_caption_url(info)
        if not cap_url:
            log(f'[{i}/{len(urls)}] 한국어 자막 없음: {info.get("title")}')
            failed += 1
            continue
        try:
            events = fetch_events(cap_url)
        except Exception as exc:                                  # noqa: BLE001
            log(f'[{i}/{len(urls)}] 자막 다운로드 실패: {exc}')
            failed += 1
            continue

        info['video_id'] = info['id']
        note = build_note(info, group_events(events, args.block_sec),
                          parse_chapters(info), created)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(note)
        log(f'[{i}/{len(urls)}] 저장 ({len(note):,}자): {os.path.basename(path)}')
        ok += 1
        time.sleep(args.sleep)

    log(f'\n완료 — 저장 {ok} / 건너뜀 {skipped} / 실패 {failed}')


if __name__ == '__main__':
    main()
