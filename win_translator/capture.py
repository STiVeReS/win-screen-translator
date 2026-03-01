import io
from dataclasses import dataclass
from typing import List, Tuple

from mss import mss
from PIL import Image

from .win32_window import get_window_rect, is_window_valid


# OCR.space free tier: 1MB. Тримаємо трохи запасу.
DEFAULT_MAX_OCR_BYTES = 950 * 1024


@dataclass
class CaptureInfo:
    left: int
    top: int
    width: int
    height: int

    # Розмір зображення, яке реально пішло в OCR (може бути з даунскейлом)
    ocr_width: int
    ocr_height: int

    # Масштаб для переведення OCR-координат у координати оверлею
    scale_x: float
    scale_y: float

    # 0 = усі монітори, 1..N = конкретний монітор
    monitor_index: int


def list_monitors() -> List[dict]:
    """Повертає список моніторів у порядку MSS (1..N)."""
    out: List[dict] = []
    with mss() as sct:
        # sct.monitors[0] = virtual screen
        idx = 1
        while idx < len(sct.monitors):
            mon = sct.monitors[idx]
            out.append({
                "index": idx,
                "left": int(mon.get("left", 0)),
                "top": int(mon.get("top", 0)),
                "width": int(mon.get("width", 0)),
                "height": int(mon.get("height", 0)),
            })
            idx += 1
    return out


def _encode_jpeg_fit(img: Image.Image, max_bytes: int) -> Tuple[bytes, int, int]:
    """Кодує картинку в JPEG так, щоб влізло в max_bytes.

    Повертає (jpeg_bytes, out_w, out_h). Якщо довелося зменшувати розмір,
    out_w/out_h відрізнятимуться від оригіналу.
    """
    if img.mode != 'RGB':
        img = img.convert('RGB')

    orig_w, orig_h = img.size

    quality = 85
    min_quality = 35
    scale = 1.0
    min_scale = 0.45

    last_bytes = b""
    last_w = orig_w
    last_h = orig_h

    while True:
        work = img
        if scale < 1.0:
            new_w = int(orig_w * scale)
            new_h = int(orig_h * scale)
            if new_w < 16:
                new_w = 16
            if new_h < 16:
                new_h = 16
            work = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        else:
            new_w = orig_w
            new_h = orig_h

        buf = io.BytesIO()
        work.save(buf, format='JPEG', quality=int(quality), optimize=True)
        data = buf.getvalue()

        last_bytes = data
        last_w = new_w
        last_h = new_h

        if len(data) <= max_bytes:
            return data, new_w, new_h

        # Спочатку знижуємо якість, потім зменшуємо розмір.
        if quality > min_quality:
            quality -= 10
            continue

        if scale > min_scale:
            # Зменшилися на 15% і пробуємо знову з нормальної якості
            scale *= 0.85
            quality = 85
            continue

        # Далі вже нема куди. Віддамо найкраще, що маємо.
        return last_bytes, last_w, last_h


def capture_screen_for_ocr(monitor_index: int = 0, max_bytes: int = DEFAULT_MAX_OCR_BYTES) -> Tuple[bytes, CaptureInfo]:
    """Знімає скріншот (virtual screen або конкретний монітор) і повертає JPEG bytes + CaptureInfo."""
    with mss() as sct:
        mons = sct.monitors

        # Валідатор індексу
        if monitor_index is None:
            monitor_index = 0

        if monitor_index < 0:
            monitor_index = 0

        if monitor_index >= len(mons):
            # Якщо користувач змінив конфіг і моніторів стало менше
            monitor_index = 0

        mon = mons[monitor_index]
        grab = sct.grab(mon)

        img = Image.frombytes('RGB', grab.size, grab.rgb)
        orig_w, orig_h = img.size

        jpeg_bytes, ocr_w, ocr_h = _encode_jpeg_fit(img, max_bytes=max_bytes)

        scale_x = 1.0
        scale_y = 1.0
        if ocr_w > 0 and ocr_h > 0:
            scale_x = float(orig_w) / float(ocr_w)
            scale_y = float(orig_h) / float(ocr_h)

        info = CaptureInfo(
            left=int(mon.get('left', 0)),
            top=int(mon.get('top', 0)),
            width=int(mon.get('width', orig_w)),
            height=int(mon.get('height', orig_h)),
            ocr_width=int(ocr_w),
            ocr_height=int(ocr_h),
            scale_x=scale_x,
            scale_y=scale_y,
            monitor_index=int(monitor_index),
        )

        return jpeg_bytes, info


def capture_screen_png(monitor_index: int = 0) -> Tuple[bytes, CaptureInfo]:
    """Знімає скріншот (virtual screen або конкретний монітор) і повертає PNG bytes + CaptureInfo.

    Для локальних OCR (RapidOCR) краще не стискати JPEG-ом: дрібний текст гине першим.
    """
    with mss() as sct:
        mons = sct.monitors

        if monitor_index is None:
            monitor_index = 0
        if monitor_index < 0:
            monitor_index = 0
        if monitor_index >= len(mons):
            monitor_index = 0

        mon = mons[monitor_index]
        grab = sct.grab(mon)
        img = Image.frombytes('RGB', grab.size, grab.rgb)
        orig_w, orig_h = img.size

        buf = io.BytesIO()
        img.save(buf, format='PNG', optimize=True)
        data = buf.getvalue()

        info = CaptureInfo(
            left=int(mon.get('left', 0)),
            top=int(mon.get('top', 0)),
            width=int(mon.get('width', orig_w)),
            height=int(mon.get('height', orig_h)),
            ocr_width=int(orig_w),
            ocr_height=int(orig_h),
            scale_x=1.0,
            scale_y=1.0,
            monitor_index=int(monitor_index),
        )

        return data, info


def _validate_monitor_index(mons, monitor_index: int) -> int:
    if monitor_index is None:
        return 0
    try:
        mi = int(monitor_index)
    except Exception:
        mi = 0
    if mi < 0:
        mi = 0
    if mi >= len(mons):
        mi = 0
    return mi


def _clamp_roi(rel_x: int, rel_y: int, width: int, height: int, mon: dict) -> Tuple[int, int, int, int]:
    mx = int(mon.get('width', 0) or 0)
    my = int(mon.get('height', 0) or 0)

    x = int(rel_x or 0)
    y = int(rel_y or 0)
    w = int(width or 1)
    h = int(height or 1)

    if x < 0:
        x = 0
    if y < 0:
        y = 0

    if w < 1:
        w = 1
    if h < 1:
        h = 1

    if mx > 0 and x >= mx:
        x = mx - 1
    if my > 0 and y >= my:
        y = my - 1

    if mx > 0 and x + w > mx:
        w = mx - x
    if my > 0 and y + h > my:
        h = my - y

    if w < 1:
        w = 1
    if h < 1:
        h = 1

    return x, y, w, h


def capture_region_for_ocr(
    monitor_index: int,
    rel_x: int,
    rel_y: int,
    width: int,
    height: int,
    max_bytes: int = DEFAULT_MAX_OCR_BYTES,
) -> Tuple[bytes, CaptureInfo]:
    """Знімає ROI з конкретного монітора і повертає JPEG bytes + CaptureInfo.

    rel_x/rel_y/width/height задаються відносно монітора (у фізичних пікселях MSS).
    """
    with mss() as sct:
        mons = sct.monitors
        mi = _validate_monitor_index(mons, monitor_index)
        if mi == 0:
            # ROI має сенс лише для конкретного монітора.
            mi = 1
            if mi >= len(mons):
                mi = 0

        mon = mons[mi]
        x, y, w, h = _clamp_roi(rel_x, rel_y, width, height, mon)

        region = {
            'left': int(mon.get('left', 0)) + int(x),
            'top': int(mon.get('top', 0)) + int(y),
            'width': int(w),
            'height': int(h),
        }

        grab = sct.grab(region)
        img = Image.frombytes('RGB', grab.size, grab.rgb)
        orig_w, orig_h = img.size

        jpeg_bytes, ocr_w, ocr_h = _encode_jpeg_fit(img, max_bytes=max_bytes)

        scale_x = 1.0
        scale_y = 1.0
        if ocr_w > 0 and ocr_h > 0:
            scale_x = float(orig_w) / float(ocr_w)
            scale_y = float(orig_h) / float(ocr_h)

        info = CaptureInfo(
            left=int(region.get('left', 0)),
            top=int(region.get('top', 0)),
            width=int(region.get('width', orig_w)),
            height=int(region.get('height', orig_h)),
            ocr_width=int(ocr_w),
            ocr_height=int(ocr_h),
            scale_x=scale_x,
            scale_y=scale_y,
            monitor_index=int(mi),
        )

        return jpeg_bytes, info


def capture_region_png(
    monitor_index: int,
    rel_x: int,
    rel_y: int,
    width: int,
    height: int,
) -> Tuple[bytes, CaptureInfo]:
    """Знімає ROI з конкретного монітора і повертає PNG bytes + CaptureInfo."""
    with mss() as sct:
        mons = sct.monitors
        mi = _validate_monitor_index(mons, monitor_index)
        if mi == 0:
            mi = 1
            if mi >= len(mons):
                mi = 0

        mon = mons[mi]
        x, y, w, h = _clamp_roi(rel_x, rel_y, width, height, mon)

        region = {
            'left': int(mon.get('left', 0)) + int(x),
            'top': int(mon.get('top', 0)) + int(y),
            'width': int(w),
            'height': int(h),
        }

        grab = sct.grab(region)
        img = Image.frombytes('RGB', grab.size, grab.rgb)
        orig_w, orig_h = img.size

        buf = io.BytesIO()
        img.save(buf, format='PNG', optimize=True)
        data = buf.getvalue()

        info = CaptureInfo(
            left=int(region.get('left', 0)),
            top=int(region.get('top', 0)),
            width=int(region.get('width', orig_w)),
            height=int(region.get('height', orig_h)),
            ocr_width=int(orig_w),
            ocr_height=int(orig_h),
            scale_x=1.0,
            scale_y=1.0,
            monitor_index=int(mi),
        )

        return data, info


def capture_target_window_png(hwnd: int) -> Tuple[bytes, CaptureInfo]:
    """Знімає PNG лише з target window (HWND) і повертає PNG bytes + CaptureInfo."""
    if not is_window_valid(hwnd):
        raise RuntimeError('Invalid target window')

    rc = get_window_rect(int(hwnd))
    if rc is None:
        raise RuntimeError('Failed to get target window rect')

    left = int(rc.left)
    top = int(rc.top)
    w = int(rc.width)
    h = int(rc.height)
    if w < 2 or h < 2:
        raise RuntimeError('Target window has invalid size')

    region = {
        'left': left,
        'top': top,
        'width': w,
        'height': h,
    }

    with mss() as sct:
        grab = sct.grab(region)
        img = Image.frombytes('RGB', grab.size, grab.rgb)
        orig_w, orig_h = img.size

        buf = io.BytesIO()
        img.save(buf, format='PNG', optimize=True)
        data = buf.getvalue()

        info = CaptureInfo(
            left=int(left),
            top=int(top),
            width=int(orig_w),
            height=int(orig_h),
            ocr_width=int(orig_w),
            ocr_height=int(orig_h),
            scale_x=1.0,
            scale_y=1.0,
            monitor_index=-1,
        )

        return data, info


def capture_target_window_for_ocr(hwnd: int, max_bytes: int = DEFAULT_MAX_OCR_BYTES) -> Tuple[bytes, CaptureInfo]:
    """Знімає target window (HWND) і повертає JPEG bytes + CaptureInfo для OCR.space."""
    if not is_window_valid(hwnd):
        raise RuntimeError('Invalid target window')

    rc = get_window_rect(int(hwnd))
    if rc is None:
        raise RuntimeError('Failed to get target window rect')

    left = int(rc.left)
    top = int(rc.top)
    w = int(rc.width)
    h = int(rc.height)
    if w < 2 or h < 2:
        raise RuntimeError('Target window has invalid size')

    region = {
        'left': left,
        'top': top,
        'width': w,
        'height': h,
    }

    with mss() as sct:
        grab = sct.grab(region)
        img = Image.frombytes('RGB', grab.size, grab.rgb)
        orig_w, orig_h = img.size

        jpeg_bytes, ocr_w, ocr_h = _encode_jpeg_fit(img, max_bytes=max_bytes)

        scale_x = 1.0
        scale_y = 1.0
        if ocr_w > 0 and ocr_h > 0:
            scale_x = float(orig_w) / float(ocr_w)
            scale_y = float(orig_h) / float(ocr_h)

        info = CaptureInfo(
            left=int(left),
            top=int(top),
            width=int(orig_w),
            height=int(orig_h),
            ocr_width=int(ocr_w),
            ocr_height=int(ocr_h),
            scale_x=scale_x,
            scale_y=scale_y,
            monitor_index=-1,
        )

        return jpeg_bytes, info
