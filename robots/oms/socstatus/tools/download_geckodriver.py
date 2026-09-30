#!/usr/bin/env python
"""Скачать geckodriver.exe (Windows x64) в корень пакета."""

from __future__ import annotations

import io
import json
import ssl
import sys
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TARGET = ROOT / "geckodriver.exe"
API = "https://api.github.com/repos/mozilla/geckodriver/releases/latest"


def main() -> int:
    if TARGET.is_file():
        print(f"Уже есть: {TARGET}")
        return 0

    ctx = ssl.create_default_context()
    print("Запрос latest release mozilla/geckodriver…")
    with urllib.request.urlopen(API, context=ctx, timeout=60) as resp:
        meta = json.loads(resp.read().decode("utf-8"))

    asset_url = None
    for asset in meta.get("assets") or []:
        name = str(asset.get("name") or "")
        if "win64" in name.lower() and name.endswith(".zip"):
            asset_url = asset.get("browser_download_url")
            break
    if not asset_url:
        print("В релизе нет win64.zip", file=sys.stderr)
        return 1

    print(f"Скачиваю {asset_url}")
    with urllib.request.urlopen(asset_url, context=ctx, timeout=180) as resp:
        data = resp.read()

    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for info in zf.infolist():
            if info.filename.lower().endswith("geckodriver.exe"):
                TARGET.write_bytes(zf.read(info))
                print(f"Сохранено: {TARGET} ({TARGET.stat().st_size} bytes)")
                return 0

    print("В архиве нет geckodriver.exe", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
