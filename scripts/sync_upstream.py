"""上游同步:抓 h9dh.cn 文件清单,比对版本,变化时下载 xlsx。

数据源是首页内嵌的 window.__INITIAL_DATA__(几 KB,自带每个文件的
版本号/更新时间/大小),无变化时只有这一次请求,对上游站零打扰。

用法:
  python sync_upstream.py [--force] [--check-only]

行为:
  - 与 data/upstream-manifest.json 比对(无 manifest 视为有变化)
  - 无变化:打印后退出 0
  - 有变化(或 --force):下载全部 xlsx 到 upstream/(文件名保留
    原始版本号),校验 zip 魔数,然后写新 manifest
  - --check-only:只比对不下载
  - 在 GitHub Actions 内(存在 GITHUB_OUTPUT)额外写 changed=true/false

manifest 的持久化策略:本脚本直接写文件,但 CI 只在验收全绿后才
commit —— 验证失败时改动随 workspace 丢弃,下次调度会重试。
"""

from __future__ import annotations

import json
import os
import sys
import time
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).parents[1]
MANIFEST = ROOT / "data" / "upstream-manifest.json"
UPSTREAM = ROOT / "upstream"
BASE = "https://www.h9dh.cn/"
UA = "TemperSync/1.0 (open-source calculator pipeline; contact via repo issues)"

# 比对用的稳定字段;id 是站点内部值,不参与比对
FIELDS = ("name", "version", "size", "updateDate", "updateTime", "url")


def fetch(url: str, timeout: int = 60) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    last_err = None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except Exception as e:  # noqa: BLE001 - 重试后统一抛出
            last_err = e
            time.sleep(5 * (attempt + 1))
    raise RuntimeError(f"下载失败 {url}: {last_err}")


def fetch_listing() -> list[dict]:
    html = fetch(BASE).decode("utf-8", errors="replace")
    marker = "__INITIAL_DATA__"
    at = html.find(marker)
    if at < 0:
        raise RuntimeError("首页里找不到 __INITIAL_DATA__,站点结构变了")
    brace = html.find("{", at)
    data, _ = json.JSONDecoder().raw_decode(html[brace:])
    files = [
        {k: f.get(k) for k in FIELDS}
        for f in data.get("files", [])
        if f.get("fileType") == "xlsx" and not f.get("isExternal")
    ]
    if not files:
        raise RuntimeError("清单里没有 xlsx 条目,站点结构变了")
    return sorted(files, key=lambda f: f["name"])


def load_manifest() -> list[dict] | None:
    if not MANIFEST.exists():
        return None
    return json.loads(MANIFEST.read_text(encoding="utf-8")).get("files")


def emit_output(changed: bool) -> None:
    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a", encoding="utf-8") as f:
            f.write(f"changed={'true' if changed else 'false'}\n")


def download_all(files: list[dict]) -> None:
    UPSTREAM.mkdir(exist_ok=True)
    for f in files:
        url = BASE + f["url"]
        dest = UPSTREAM / f["name"]
        blob = fetch(url, timeout=120)
        if blob[:2] != b"PK":
            raise RuntimeError(
                f"{f['name']} 不是 xlsx(前 2 字节 {blob[:2]!r}),"
                "可能被反爬页面拦截"
            )
        dest.write_bytes(blob)
        print(f"  下载 {f['name']} ({len(blob) / 1024:.0f} KB)")


def main() -> None:
    force = "--force" in sys.argv
    check_only = "--check-only" in sys.argv

    remote = fetch_listing()
    local = load_manifest()
    changed = force or local is None or remote != local

    if not changed:
        print(f"无变化({len(remote)} 个文件,与 manifest 一致)")
        emit_output(False)
        return

    diff = []
    old_by_name = {f["name"]: f for f in (local or [])}
    for f in remote:
        prev = old_by_name.get(f["name"])
        if prev is None:
            diff.append(f"新增 {f['name']}")
        elif prev != f:
            diff.append(f"更新 {f['name']} (v{prev.get('version')} -> v{f.get('version')})")
    for name in old_by_name.keys() - {f["name"] for f in remote}:
        diff.append(f"移除 {name}")
    print("检测到上游变化:" + ("(--force 全量刷新)" if force and not diff else ""))
    for line in diff:
        print(f"  {line}")

    if check_only:
        emit_output(True)
        return

    download_all(remote)
    MANIFEST.write_text(
        json.dumps(
            {
                "_meta": {
                    "source": BASE,
                    "syncedAt": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                },
                "files": remote,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    print(f"manifest 已更新: {MANIFEST.relative_to(ROOT)}")
    emit_output(True)


if __name__ == "__main__":
    main()
