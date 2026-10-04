"""上游同步:抓 h9dh.cn 文件清单,比对版本,变化时下载 xlsx。

数据源是 /api/home JSON 接口(几 KB,自带每个文件的版本号/更新时间/
大小),无变化时只有这一次请求,对上游站零打扰。
2026-10 起站点改为 SPA,原首页内嵌 window.__INITIAL_DATA__ 已下线。

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

# 比对用的稳定字段;id/category 等站点内部值不参与比对
FIELDS = ("name", "version", "size", "updatedAt", "url")


def fetch(url: str, timeout: int = 60, want_zip: bool = False) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    last_err = None
    for attempt in range(5):  # CI 出口到上游偶发读超时,多试几次
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                blob = resp.read()
                if want_zip and blob[:2] != b"PK":
                    # 上游偶发对 CI 出口 IP 弹反爬页(HTTP 200 + HTML),
                    # 与网络异常一样按可重试处理
                    raise IOError(f"返回的不是 xlsx(前 2 字节 {blob[:2]!r})")
                return blob
        except Exception as e:  # noqa: BLE001 - 重试后统一抛出
            last_err = e
            time.sleep(10 * (attempt + 1))
    raise RuntimeError(f"下载失败 {url}: {last_err}")


def fetch_listing() -> list[dict]:
    blob = fetch(BASE + "api/home")
    try:
        data = json.loads(blob.decode("utf-8"))
    except json.JSONDecodeError as e:
        raise RuntimeError(f"/api/home 返回的不是 JSON,站点结构变了: {e}") from e
    files = [
        {
            "name": f.get("name"),
            "version": f.get("version"),
            "size": f.get("size_bytes"),
            "updatedAt": f.get("updated_at"),
            "url": f.get("download_url"),
        }
        for f in data.get("files", [])
        if f.get("kind") == "file"
        and f.get("extension") == "xlsx"
        and f.get("visible")
        and not f.get("external_url")
        and not f.get("deleted_at")
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
    for i, f in enumerate(files):
        # 旧站点 url 不带前导斜杠,新接口带;必须归一化,
        # 双斜杠会被上游 WAF 拦成 HTML 反爬页
        url = BASE.rstrip("/") + "/" + f["url"].lstrip("/")
        dest = UPSTREAM / f["name"]
        blob = fetch(url, timeout=120, want_zip=True)
        if i > 0:
            time.sleep(1)  # 连续下载间隔,降低触发上游反爬的概率
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
