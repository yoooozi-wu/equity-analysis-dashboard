# -*- coding: utf-8 -*-
"""下载 ECharts 5.4.4 到 assets/echarts.min.js。

看板构建需要把 ECharts 内嵌进 HTML（离线自包含），因此首次使用前先运行本脚本拉取一次。
用法：
    python scripts/fetch_echarts.py
"""
import urllib.request
from pathlib import Path

URL = "https://cdn.jsdelivr.net/npm/echarts@5.4.4/dist/echarts.min.js"
OUT = Path(__file__).resolve().parent.parent / "assets" / "echarts.min.js"


def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    print("下载:", URL)
    with urllib.request.urlopen(URL, timeout=60) as r:
        data = r.read()
    OUT.write_bytes(data)
    print("已保存:", OUT, "(%.2f MB)" % (len(data) / 1024 / 1024))


if __name__ == "__main__":
    main()
