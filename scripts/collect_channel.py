#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""무꿈사TV 채널 전수 수집기.

YouTube 공개 페이지의 ytInitialData / innertube browse·player 엔드포인트만 사용.
API 키·로그인 불필요. 결과는 out/ 에 JSON으로 떨어진다.

    python3 scripts/collect_channel.py --out out

주의: 워치 페이지(watch?v=)는 공유 IP에서 429가 잦아 자막(transcript)은 수집하지 않는다.
      영상 설명(description)은 innertube player 엔드포인트로 안정적으로 수집된다.
"""
import argparse, json, os, re, subprocess, sys, time, urllib.request

CHANNEL_URL = "https://www.youtube.com/@%EB%AC%B4%EA%BF%88%EC%82%AC"
CHANNEL_ID  = "UCjUySF5j9cSNlsR2G6sW3XQ"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")
TABS = {  # 채널 탭별 browse params
    "videos":  "EgZ2aWRlb3PyBgQKAjoA",
    "streams": "EgdzdHJlYW1z8gYECgJ6AA==",
}


def curl(url: str) -> str:
    r = subprocess.run(["curl", "-sL", "--max-time", "60", "-A", UA,
                        "-H", "Accept-Language: ko-KR,ko;q=0.9", url],
                       capture_output=True, text=True)
    return r.stdout


def initial_data(html: str) -> dict:
    m = re.search(r"var ytInitialData = (\{.*?\});</script>", html, re.S)
    if not m:
        raise SystemExit("ytInitialData 를 찾지 못했다. 페이지 구조가 바뀌었을 수 있음.")
    return json.loads(m.group(1))


def innertube_config(html: str):
    key = re.search(r'"INNERTUBE_API_KEY":"([^"]+)"', html).group(1)
    ctx = json.loads(re.search(r'"INNERTUBE_CONTEXT":(\{.+?\}),"INNERTUBE_CONTEXT_CLIENT_NAME"',
                               html, re.S).group(1))
    return key, ctx


def post(endpoint: str, key: str, ctx: dict, body: dict, retries: int = 4) -> dict:
    url = f"https://www.youtube.com/youtubei/v1/{endpoint}?key={key}&prettyPrint=false"
    req = urllib.request.Request(url, data=json.dumps(body).encode(), headers={
        "Content-Type": "application/json", "User-Agent": UA,
        "Accept-Language": "ko-KR,ko;q=0.9", "Origin": "https://www.youtube.com",
        "X-Youtube-Client-Name": "1",
        "X-Youtube-Client-Version": ctx["client"]["clientVersion"]})
    for attempt in range(retries):
        try:
            return json.loads(urllib.request.urlopen(req, timeout=60).read())
        except Exception as exc:                       # noqa: BLE001
            print(f"  retry {attempt}: {exc}", file=sys.stderr)
            time.sleep(2 ** attempt)
    return {}


def parse_lockup(lv: dict) -> dict:
    """지금의 YouTube 는 videoRenderer 대신 lockupViewModel 을 쓴다."""
    md = lv.get("metadata", {}).get("lockupMetadataViewModel", {})
    rows = [p["text"]["content"]
            for r in md.get("metadata", {}).get("contentMetadataViewModel", {}).get("metadataRows", [])
            for p in r.get("metadataParts", []) if p.get("text", {}).get("content")]
    duration = None
    for ov in lv.get("contentImage", {}).get("thumbnailViewModel", {}).get("overlays", []):
        for badge in ov.get("thumbnailBottomOverlayViewModel", {}).get("badges", []):
            text = badge.get("thumbnailBadgeViewModel", {}).get("text")
            if text and re.fullmatch(r"\d+(:\d{2})+", text):
                duration = text
    return {"id": lv.get("contentId"), "title": (md.get("title") or {}).get("content"),
            "duration": duration, "meta": rows,
            "kind": lv.get("contentType")}


def walk(node, videos: list, continuations: list) -> None:
    if isinstance(node, dict):
        lv = node.get("lockupViewModel")
        if isinstance(lv, dict):
            videos.append(parse_lockup(lv))
        cont = node.get("continuationItemRenderer")
        if cont:
            token = cont.get("continuationEndpoint", {}).get("continuationCommand", {}).get("token")
            if token:
                continuations.append(token)
        for value in node.values():
            walk(value, videos, continuations)
    elif isinstance(node, list):
        for value in node:
            walk(value, videos, continuations)


def paginate(seed: dict, key: str, ctx: dict, limit: int = 100) -> list:
    out, seen = [], set()
    videos, conts = [], []
    walk(seed, videos, conts)
    for v in videos:
        if v["id"] and v["id"] not in seen:
            seen.add(v["id"]); out.append(v)
    token = conts[0] if conts else None
    for _ in range(limit):
        if not token:
            break
        page = post("browse", key, ctx, {"context": ctx, "continuation": token})
        videos, conts = [], []
        walk(page, videos, conts)
        fresh = 0
        for v in videos:
            if v["id"] and v["id"] not in seen:
                seen.add(v["id"]); out.append(v); fresh += 1
        token = conts[0] if conts else None
        if not fresh:
            break
        time.sleep(0.4)
    return out


def collect_videos(key, ctx):
    found = {}
    for name, params in TABS.items():
        seed = post("browse", key, ctx, {"context": ctx, "browseId": CHANNEL_ID,
                                         "params": params.replace("%3D", "=")})
        items = paginate(seed, key, ctx)
        print(f"  tab {name}: {len(items)}", file=sys.stderr)
        for v in items:
            found.setdefault(v["id"], v)
    return found


def collect_playlists(key, ctx):
    html = curl(CHANNEL_URL + "/playlists")
    videos, _ = [], []
    walk(initial_data(html), videos, _)
    playlists = [v for v in videos if (v["id"] or "").startswith("PL")]
    out = {}
    for pl in playlists:
        page = initial_data(curl(f"https://www.youtube.com/playlist?list={pl['id']}"))
        items, conts = [], []
        walk(page, items, conts)
        out[pl["id"]] = {"title": pl["title"],
                         "video_ids": [i["id"] for i in items
                                       if i.get("kind") == "LOCKUP_CONTENT_TYPE_VIDEO"]}
        print(f"  playlist {pl['title'][:30]}: {len(out[pl['id']]['video_ids'])}", file=sys.stderr)
        time.sleep(0.8)
    return out


def collect_details(video_ids, key, ctx):
    """설명·조회수·게시일. 자막은 봇 차단으로 수집 불가."""
    out = {}
    for i, vid in enumerate(video_ids):
        data = post("player", key, ctx, {"context": ctx, "videoId": vid,
                                         "contentCheckOk": True, "racyCheckOk": True})
        details = data.get("videoDetails", {})
        micro = data.get("microformat", {}).get("playerMicroformatRenderer", {})
        out[vid] = {"title": details.get("title"),
                    "description": details.get("shortDescription", ""),
                    "keywords": details.get("keywords", []),
                    "views": details.get("viewCount"),
                    "length_sec": details.get("lengthSeconds"),
                    "published": micro.get("publishDate")}
        if i % 25 == 0:
            print(f"  details {i}/{len(video_ids)}", file=sys.stderr)
        time.sleep(0.25)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="out")
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)

    html = curl(CHANNEL_URL + "/videos")
    key, ctx = innertube_config(html)

    print("collecting videos...", file=sys.stderr)
    videos = collect_videos(key, ctx)
    print("collecting playlists...", file=sys.stderr)
    playlists = collect_playlists(key, ctx)
    for pl in playlists.values():                    # 탭에 안 잡히는 영상 보강
        for vid in pl["video_ids"]:
            videos.setdefault(vid, {"id": vid, "title": None, "duration": None, "meta": []})

    print(f"collecting details for {len(videos)} videos...", file=sys.stderr)
    details = collect_details(sorted(videos), key, ctx)

    for path, payload in [("videos_raw.json", videos),
                          ("playlists.json", playlists),
                          ("details.json", details)]:
        with open(os.path.join(args.out, path), "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=1)
    print(f"done: {len(videos)} videos -> {args.out}/", file=sys.stderr)


if __name__ == "__main__":
    main()
