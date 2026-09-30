from __future__ import annotations

from pathlib import Path

import pytesseract
from PIL import Image, ImageOps

from browser_auto.config import CAPTCHA_IMAGE_DIR, TESSERACT_CMD

if TESSERACT_CMD.is_file():
    pytesseract.pytesseract.tesseract_cmd = str(TESSERACT_CMD)


def ensure_tesseract() -> Path:
    if not TESSERACT_CMD.is_file():
        raise FileNotFoundError(
            f"Tesseract не найден: {TESSERACT_CMD}. "
            "Установите UB-Mannheim.TesseractOCR или укажите TESSERACT_CMD в .env"
        )
    return TESSERACT_CMD


def _trim_light_borders(img: Image.Image, *, light_threshold: int = 160, pad: int = 1) -> Image.Image:
    """
    Убрать светлую рамку вокруг капчи ИСЗЛ (часто 1–10px справа/снизу от screenshot элемента).
    Оставляем тёмную область с цифрами + небольшой pad.
    """
    gray = ImageOps.grayscale(img)
    w, h = gray.size
    px = gray.load()
    # строки/столбцы, где есть достаточно тёмных пикселей (фон капчи)
    dark_cols = [x for x in range(w) if sum(1 for y in range(h) if px[x, y] < light_threshold) >= max(2, h // 5)]
    dark_rows = [y for y in range(h) if sum(1 for x in range(w) if px[x, y] < light_threshold) >= max(2, w // 5)]
    if not dark_cols or not dark_rows:
        return img
    left = max(0, dark_cols[0] - pad)
    right = min(w, dark_cols[-1] + 1 + pad)
    top = max(0, dark_rows[0] - pad)
    bottom = min(h, dark_rows[-1] + 1 + pad)
    if right - left < 8 or bottom - top < 8:
        return img
    if (left, top, right, bottom) == (0, 0, w, h):
        return img
    return img.crop((left, top, right, bottom))


def recognize_digits(
    image_path: Path | str,
    *,
    crop_box: tuple[int, int, int, int] | None = None,
    expected_len: int = 4,
    auto_trim: bool = True,
) -> str:
    """OCR цифровой капчи (ИСЗЛ и др.)."""
    ensure_tesseract()

    img = Image.open(image_path)
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGB")
    if crop_box:
        img = img.crop(crop_box)
    if auto_trim:
        img = _trim_light_borders(img)

    gray = ImageOps.grayscale(img)
    # белые цифры на чёрном → чёрные на белом
    if sum(gray.getdata()) / max(1, gray.width * gray.height) < 128:
        gray = ImageOps.invert(gray)

    # мелкие iframe-капчи (~70x22) tesseract читает только после увеличения
    if gray.width < 160 or gray.height < 40:
        gray = gray.resize((gray.width * 4, gray.height * 4), Image.Resampling.LANCZOS)

    bw = gray.point(lambda x: 0 if x < 140 else 255)

    best = ""
    for image in (bw, gray):
        for psm in (7, 8, 13):
            config = rf"--oem 3 --psm {psm} -c tessedit_char_whitelist=0123456789"
            text = pytesseract.image_to_string(image, config=config)
            digits = "".join(ch for ch in text if ch.isdigit())
            if expected_len and len(digits) == expected_len:
                return digits
            if len(digits) > len(best):
                best = digits
    return best

def captcha_temp_path(name: str = "temp.png") -> Path:
    return CAPTCHA_IMAGE_DIR / name
