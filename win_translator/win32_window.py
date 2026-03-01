import os
from dataclasses import dataclass
from typing import Optional, Tuple


@dataclass
class WinRect:
    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self) -> int:
        return max(0, int(self.right) - int(self.left))

    @property
    def height(self) -> int:
        return max(0, int(self.bottom) - int(self.top))


if os.name != 'nt':
    def is_window_valid(hwnd: int) -> bool:
        return False

    def get_window_title(hwnd: int) -> str:
        return ''

    def get_window_under_cursor(exclude_hwnd: Optional[int] = None) -> Tuple[int, str]:
        return 0, ''

    def get_window_rect(hwnd: int, client_only: bool = True) -> Optional[WinRect]:
        return None

    def capture_window_bgra(hwnd: int, client_only: bool = True) -> Tuple[bytes, Optional[WinRect]]:
        return b'', None

else:
    import ctypes
    import ctypes.wintypes

    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32

    PW_CLIENTONLY = 0x00000001
    PW_RENDERFULLCONTENT = 0x00000002
    DIB_RGB_COLORS = 0
    SRCCOPY = 0x00CC0020

    GA_ROOT = 2


    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.wintypes.LONG), ("y", ctypes.wintypes.LONG)]


    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.wintypes.LONG),
            ("top", ctypes.wintypes.LONG),
            ("right", ctypes.wintypes.LONG),
            ("bottom", ctypes.wintypes.LONG),
        ]


    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", ctypes.wintypes.DWORD),
            ("biWidth", ctypes.wintypes.LONG),
            ("biHeight", ctypes.wintypes.LONG),
            ("biPlanes", ctypes.wintypes.WORD),
            ("biBitCount", ctypes.wintypes.WORD),
            ("biCompression", ctypes.wintypes.DWORD),
            ("biSizeImage", ctypes.wintypes.DWORD),
            ("biXPelsPerMeter", ctypes.wintypes.LONG),
            ("biYPelsPerMeter", ctypes.wintypes.LONG),
            ("biClrUsed", ctypes.wintypes.DWORD),
            ("biClrImportant", ctypes.wintypes.DWORD),
        ]


    class BITMAPINFO(ctypes.Structure):
        _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", ctypes.wintypes.DWORD * 3)]


    def is_window_valid(hwnd: int) -> bool:
        try:
            if hwnd is None:
                return False
            h = int(hwnd)
            if h <= 0:
                return False
            return True
        except Exception:
            return False


    def get_window_title(hwnd: int) -> str:
        try:
            h = int(hwnd)
            if h <= 0:
                return ''
            length = int(user32.GetWindowTextLengthW(ctypes.wintypes.HWND(h)) or 0)
            if length <= 0:
                return ''
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(ctypes.wintypes.HWND(h), buf, length + 1)
            return str(buf.value or '').strip()
        except Exception:
            return ''


    def get_window_under_cursor(exclude_hwnd: Optional[int] = None) -> Tuple[int, str]:
        """Повертає (hwnd, title) top-level вікна під курсором.

        exclude_hwnd: якщо переданий і збігається, повернемо (0, '').
        """
        try:
            pt = POINT()
            ok = bool(user32.GetCursorPos(ctypes.byref(pt)))
            if not ok:
                return 0, ''
            hwnd = int(user32.WindowFromPoint(pt) or 0)
            if hwnd <= 0:
                return 0, ''

            # Піднімаємось до root window
            hwnd_root = int(user32.GetAncestor(ctypes.wintypes.HWND(hwnd), GA_ROOT) or hwnd)
            if hwnd_root <= 0:
                hwnd_root = hwnd

            if exclude_hwnd is not None:
                try:
                    if int(exclude_hwnd) == int(hwnd_root):
                        return 0, ''
                except Exception:
                    pass

            if not bool(user32.IsWindowVisible(ctypes.wintypes.HWND(hwnd_root))):
                return 0, ''

            title = get_window_title(hwnd_root)
            if not title:
                # Деякі вікна без заголовка нам не цікаві.
                return 0, ''
            return hwnd_root, title
        except Exception:
            return 0, ''


    def get_window_rect(hwnd: int, client_only: bool = True) -> Optional[WinRect]:
        try:
            h = int(hwnd)
            if h <= 0:
                return None

            if client_only:
                rc = RECT()
                ok = bool(user32.GetClientRect(ctypes.wintypes.HWND(h), ctypes.byref(rc)))
                if not ok:
                    return None

                # rc у клієнтських координатах, переводимо в screen
                pt = POINT()
                pt.x = 0
                pt.y = 0
                ok2 = bool(user32.ClientToScreen(ctypes.wintypes.HWND(h), ctypes.byref(pt)))
                if not ok2:
                    return None

                left = int(pt.x)
                top = int(pt.y)
                right = int(pt.x + (rc.right - rc.left))
                bottom = int(pt.y + (rc.bottom - rc.top))
                return WinRect(left=left, top=top, right=right, bottom=bottom)

            rcw = RECT()
            ok = bool(user32.GetWindowRect(ctypes.wintypes.HWND(h), ctypes.byref(rcw)))
            if not ok:
                return None
            return WinRect(left=int(rcw.left), top=int(rcw.top), right=int(rcw.right), bottom=int(rcw.bottom))
        except Exception:
            return None


    def _grab_via_printwindow(hwnd: int, w: int, h: int, flags: int) -> Optional[bytes]:
        """Повертає BGRA bytes (w*h*4) або None."""
        hdc_screen = None
        hdc_mem = None
        hbmp = None
        old_obj = None
        try:
            hdc_screen = user32.GetDC(0)
            if not hdc_screen:
                return None

            hdc_mem = gdi32.CreateCompatibleDC(hdc_screen)
            if not hdc_mem:
                return None

            hbmp = gdi32.CreateCompatibleBitmap(hdc_screen, int(w), int(h))
            if not hbmp:
                return None

            old_obj = gdi32.SelectObject(hdc_mem, hbmp)

            ok = bool(user32.PrintWindow(ctypes.wintypes.HWND(int(hwnd)), hdc_mem, int(flags)))
            if not ok:
                return None

            bmi = BITMAPINFO()
            bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bmi.bmiHeader.biWidth = int(w)
            # негативна висота = top-down DIB
            bmi.bmiHeader.biHeight = -int(h)
            bmi.bmiHeader.biPlanes = 1
            bmi.bmiHeader.biBitCount = 32
            bmi.bmiHeader.biCompression = 0  # BI_RGB
            bmi.bmiHeader.biSizeImage = int(w) * int(h) * 4

            buf = ctypes.create_string_buffer(int(w) * int(h) * 4)
            got = gdi32.GetDIBits(hdc_mem, hbmp, 0, int(h), buf, ctypes.byref(bmi), DIB_RGB_COLORS)
            if int(got) == 0:
                return None

            return bytes(buf.raw)
        except Exception:
            return None
        finally:
            try:
                if old_obj is not None and hdc_mem is not None:
                    gdi32.SelectObject(hdc_mem, old_obj)
            except Exception:
                pass
            try:
                if hbmp is not None:
                    gdi32.DeleteObject(hbmp)
            except Exception:
                pass
            try:
                if hdc_mem is not None:
                    gdi32.DeleteDC(hdc_mem)
            except Exception:
                pass
            try:
                if hdc_screen is not None:
                    user32.ReleaseDC(0, hdc_screen)
            except Exception:
                pass


    def _grab_via_bitblt(rect: WinRect) -> Optional[bytes]:
        """Фолбек: знімаємо прямокутник зі скріну (BGRA)."""
        hdc_screen = None
        hdc_mem = None
        hbmp = None
        old_obj = None
        try:
            w = int(rect.width)
            h = int(rect.height)
            if w <= 0 or h <= 0:
                return None

            hdc_screen = user32.GetDC(0)
            if not hdc_screen:
                return None

            hdc_mem = gdi32.CreateCompatibleDC(hdc_screen)
            if not hdc_mem:
                return None

            hbmp = gdi32.CreateCompatibleBitmap(hdc_screen, int(w), int(h))
            if not hbmp:
                return None

            old_obj = gdi32.SelectObject(hdc_mem, hbmp)
            ok = bool(gdi32.BitBlt(hdc_mem, 0, 0, int(w), int(h), hdc_screen, int(rect.left), int(rect.top), SRCCOPY))
            if not ok:
                return None

            bmi = BITMAPINFO()
            bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bmi.bmiHeader.biWidth = int(w)
            bmi.bmiHeader.biHeight = -int(h)
            bmi.bmiHeader.biPlanes = 1
            bmi.bmiHeader.biBitCount = 32
            bmi.bmiHeader.biCompression = 0
            bmi.bmiHeader.biSizeImage = int(w) * int(h) * 4

            buf = ctypes.create_string_buffer(int(w) * int(h) * 4)
            got = gdi32.GetDIBits(hdc_mem, hbmp, 0, int(h), buf, ctypes.byref(bmi), DIB_RGB_COLORS)
            if int(got) == 0:
                return None
            return bytes(buf.raw)
        except Exception:
            return None
        finally:
            try:
                if old_obj is not None and hdc_mem is not None:
                    gdi32.SelectObject(hdc_mem, old_obj)
            except Exception:
                pass
            try:
                if hbmp is not None:
                    gdi32.DeleteObject(hbmp)
            except Exception:
                pass
            try:
                if hdc_mem is not None:
                    gdi32.DeleteDC(hdc_mem)
            except Exception:
                pass
            try:
                if hdc_screen is not None:
                    user32.ReleaseDC(0, hdc_screen)
            except Exception:
                pass


    def capture_window_bgra(hwnd: int, client_only: bool = True) -> Tuple[bytes, Optional[WinRect]]:
        """Захоплення вікна (BGRA bytes) без сторонніх оверлеїв.

        1) пробуємо PrintWindow (не захоплює чужі вікна поверх)
        2) фолбек на BitBlt з екрана (може захопити накладення, якщо є)
        """
        if not is_window_valid(hwnd):
            return b'', None

        rect = get_window_rect(hwnd, client_only=client_only)
        if rect is None:
            return b'', None

        w = int(rect.width)
        h = int(rect.height)
        if w <= 0 or h <= 0:
            return b'', None

        flags = PW_RENDERFULLCONTENT
        if client_only:
            flags |= PW_CLIENTONLY

        data = _grab_via_printwindow(hwnd, w, h, flags)
        if data is None:
            data = _grab_via_bitblt(rect)
        if data is None:
            return b'', rect

        return data, rect
