from __future__ import annotations

from dataclasses import dataclass, field
import textwrap
from typing import List, Optional, Sequence

from PySide6 import QtCore, QtGui, QtWidgets


def _force_topmost(win_id: int) -> None:
    """Пробуємо примусово зробити вікно TOPMOST на Windows.

    Не вирішує ексклюзивний fullscreen у частини ігор, але для windowed/borderless часто допомагає.
    """
    try:
        import sys
        if sys.platform != 'win32':
            return
        import ctypes

        HWND_TOPMOST = -1
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_NOACTIVATE = 0x0010
        SWP_SHOWWINDOW = 0x0040
        ctypes.windll.user32.SetWindowPos(
            int(win_id),
            HWND_TOPMOST,
            0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW,
        )
    except Exception:
        pass


def _clamp_int(v: int, lo: int, hi: int) -> int:
    try:
        out = int(v)
    except Exception:
        out = lo

    if out < lo:
        out = lo
    if out > hi:
        out = hi
    return out


def _sanitize_rgb(rgb: Optional[Sequence[int]], fallback: Sequence[int]) -> List[int]:
    if rgb is None:
        return [int(fallback[0]), int(fallback[1]), int(fallback[2])]

    try:
        if len(rgb) != 3:
            return [int(fallback[0]), int(fallback[1]), int(fallback[2])]

        r = _clamp_int(int(rgb[0]), 0, 255)
        g = _clamp_int(int(rgb[1]), 0, 255)
        b = _clamp_int(int(rgb[2]), 0, 255)
        return [r, g, b]
    except Exception:
        return [int(fallback[0]), int(fallback[1]), int(fallback[2])]


def _wrap_text_by_chars(text: str, width: int) -> str:
    w = int(width or 0)
    if w <= 0:
        return text

    parts: List[str] = []
    for p in (text or '').splitlines():
        p2 = p.strip()
        if not p2:
            parts.append('')
            continue

        parts.append(
            textwrap.fill(
                p2,
                width=w,
                break_long_words=True,
                break_on_hyphens=False,
            )
        )

    return '\n'.join(parts)


def _layout_lines_word_wrap(text: str, font: QtGui.QFont, max_width_px: int) -> List[str]:
    """Розкладає текст на рядки так, як його буде переносити Qt по ширині в пікселях.

    ВАЖЛИВО: 
 між рядками не означає "порожній рядок". Це просто перехід на новий рядок.
    Тому не вставляємо додаткові '' між параграфами, а зберігаємо тільки ті порожні рядки,
    які реально є у вихідному тексті.
    """
    w = int(max_width_px or 0)
    src = text or ''

    if w <= 0:
        lines = src.split('\n')
        return lines if lines else ['']

    option = QtGui.QTextOption()
    option.setWrapMode(QtGui.QTextOption.WrapMode.WrapAtWordBoundaryOrAnywhere)

    out: List[str] = []

    # split('\n') зберігає порожні рядки (на відміну від splitlines())
    parts = src.split('\n')
    if not parts:
        parts = ['']

    for part in parts:
        if part == '':
            out.append('')
            continue

        layout = QtGui.QTextLayout(part, font)
        layout.setTextOption(option)
        layout.beginLayout()
        while True:
            ln = layout.createLine()
            if not ln.isValid():
                break
            ln.setLineWidth(float(w))
            start = int(ln.textStart())
            length = int(ln.textLength())
            out.append(part[start:start + length])
        layout.endLayout()

    if not out:
        out = ['']
    return out


def _elide_line(fm: QtGui.QFontMetrics, line: str, max_width_px: int) -> str:
    w = int(max_width_px or 0)
    if w <= 0:
        return line
    try:
        return fm.elidedText(line, QtCore.Qt.TextElideMode.ElideRight, w)
    except Exception:
        return line


@dataclass
class OverlayItem:
    left: int
    top: int
    right: int
    bottom: int
    text: str
    bg_color: Optional[list[int]] = None

    @property
    def width(self) -> int:
        return max(1, self.right - self.left)

    @property
    def height(self) -> int:
        return max(1, self.bottom - self.top)


@dataclass
class OverlayStyle:
    font_family: str = 'Segoe UI'
    font_size: float = 14.0
    padding: int = 6
    round_radius: int = 8

    bg_color: List[int] = field(default_factory=lambda: [0, 0, 0])
    bg_opacity: int = 170

    text_color: List[int] = field(default_factory=lambda: [255, 255, 255])
    text_opacity: int = 235

    use_ocr_bg_color: bool = False
    max_chars_per_line: int = 0
    max_box_height: int = 0


class OverlayWindow(QtWidgets.QWidget):
    """Прозорий оверлей поверх усіх вікон, який малює переклад."""

    def __init__(self):
        super().__init__(None)
        self._items: List[OverlayItem] = []
        self._font_scale: float = 1.0
        self._style = OverlayStyle()

        self._exclude_from_capture: bool = False
        self._exclude_from_capture_applied: bool = False

        self.setWindowFlag(QtCore.Qt.WindowType.FramelessWindowHint, True)
        self.setWindowFlag(QtCore.Qt.WindowType.WindowStaysOnTopHint, True)
        self.setWindowFlag(QtCore.Qt.WindowType.Tool, True)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        # Трохи швидше перемальовка. Але без ручного очищення (нижче) буде «ghosting».
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_NoSystemBackground, True)

    def set_exclude_from_capture(self, enabled: bool) -> None:
        """На Windows намагається виключити overlay з будь-яких screen-capture.

        Якщо ОС не підтримує, метод тихо нічого не зробить.
        """
        self._exclude_from_capture = bool(enabled)
        self._apply_exclude_from_capture()

    def is_excluded_from_capture(self) -> bool:
        return bool(self._exclude_from_capture_applied)

    def _apply_exclude_from_capture(self) -> None:
        try:
            import sys
            if sys.platform != 'win32':
                self._exclude_from_capture_applied = False
                return
            import ctypes

            hwnd = int(self.winId() or 0)
            if hwnd <= 0:
                self._exclude_from_capture_applied = False
                return

            # Windows 10 2004+: WDA_EXCLUDEFROMCAPTURE (0x11)
            WDA_NONE = 0x0
            WDA_EXCLUDEFROMCAPTURE = 0x11

            mode = int(WDA_NONE)
            if bool(self._exclude_from_capture):
                mode = int(WDA_EXCLUDEFROMCAPTURE)

            ok = bool(ctypes.windll.user32.SetWindowDisplayAffinity(hwnd, mode))
            self._exclude_from_capture_applied = bool(ok) and bool(self._exclude_from_capture)
        except Exception:
            self._exclude_from_capture_applied = False

    def set_overlay_geometry(self, left: int, top: int, width: int, height: int) -> None:
        self.setGeometry(left, top, width, height)

    def set_items(self, items: List[OverlayItem], font_scale: float = 1.0) -> None:
        self._items = items or []
        self._font_scale = max(0.5, min(3.0, float(font_scale or 1.0)))
        self.update()

        # Підстрахуємось: ще раз піднімемо вікно наверх.
        _force_topmost(int(self.winId()))

    def set_style(self, style: OverlayStyle) -> None:
        if style is None:
            return
        self._style = style
        self.update()

    def showEvent(self, event: QtGui.QShowEvent) -> None:
        super().showEvent(event)
        _force_topmost(int(self.winId()))
        # На Windows інколи affinity застосовується стабільніше після show.
        self._apply_exclude_from_capture()

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)

        # Примусово чистимо весь прозорий буфер.
        # На layered windows (Windows) та при частковій перерисовці Qt може малювати тільки dirty-rect,
        # через що старі бокси «залипають» і накладаються один на одне.
        try:
            painter.setClipping(False)
            painter.setClipRect(self.rect(), QtCore.Qt.ClipOperation.ReplaceClip)
        except Exception:
            pass

        painter.setCompositionMode(QtGui.QPainter.CompositionMode.CompositionMode_Source)
        painter.fillRect(self.rect(), QtGui.QColor(0, 0, 0, 0))
        painter.setCompositionMode(QtGui.QPainter.CompositionMode.CompositionMode_SourceOver)

        if not self._items:
            painter.end()
            return

        style = self._style

        base_font = QtGui.QFont(style.font_family or 'Segoe UI')
        base_font.setPointSizeF(float(style.font_size or 14.0) * float(self._font_scale or 1.0))
        painter.setFont(base_font)

        fm = QtGui.QFontMetrics(base_font)
        line_h = int(fm.lineSpacing() or 0)
        if line_h <= 0:
            line_h = int(fm.height() or 0)
        if line_h <= 0:
            line_h = 14

        pad = _clamp_int(int(style.padding or 0), 0, 80)
        rr = _clamp_int(int(style.round_radius or 0), 0, 80)

        max_box_h = 0
        try:
            max_box_h = int(getattr(style, 'max_box_height', 0) or 0)
        except Exception:
            max_box_h = 0
        if max_box_h < 0:
            max_box_h = 0

        tx_rgb = _sanitize_rgb(style.text_color, [255, 255, 255])
        tx_a = _clamp_int(int(style.text_opacity or 0), 0, 255)
        text_pen = QtGui.QPen(QtGui.QColor(int(tx_rgb[0]), int(tx_rgb[1]), int(tx_rgb[2]), int(tx_a)))

        bg_rgb_default = _sanitize_rgb(style.bg_color, [0, 0, 0])
        bg_a_default = _clamp_int(int(style.bg_opacity or 0), 0, 255)

        # ВАЖЛИВО: AlignTop, інакше Qt може підрізати текст по висоті при WordWrap/HiDPI.
        flags = (
            QtCore.Qt.AlignmentFlag.AlignLeft |
            QtCore.Qt.AlignmentFlag.AlignTop |
            QtCore.Qt.TextFlag.TextWordWrap
        )

        win_w = int(self.width() or 0)
        win_h = int(self.height() or 0)
        if win_w < 1:
            win_w = 1
        if win_h < 1:
            win_h = 1

        for it in self._items:
            # it.left/top = позиція тексту (верх-ліворуч OCR rect) у координатах оверлею (DIP)
            text_x = int(it.left)
            text_y = int(it.top)
            base_w = int(it.width)
            base_h = int(it.height)

            if base_w < 1:
                base_w = 1
            if base_h < 1:
                base_h = 1

            if text_x < 0:
                text_x = 0
            if text_y < 0:
                text_y = 0

            txt = it.text or ''
            wrap_chars = int(style.max_chars_per_line or 0)
            if wrap_chars > 0:
                txt = _wrap_text_by_chars(txt, wrap_chars)

            # Внутрішня ширина (без padding). Стартуємо від OCR rect ширини.
            inner_w = base_w

            if wrap_chars > 0:
                lines_probe = (txt or '').splitlines()
                if not lines_probe:
                    lines_probe = ['']
                max_line_px = 0
                for ln in lines_probe:
                    try:
                        w = int(fm.horizontalAdvance(ln) or 0)
                    except Exception:
                        w = 0
                    if w > max_line_px:
                        max_line_px = w
                if max_line_px > inner_w:
                    inner_w = max_line_px
            else:
                try:
                    br = fm.boundingRect(QtCore.QRect(0, 0, int(base_w), 10000), flags, txt)
                    bw = int(br.width() or 0)
                    if bw > inner_w:
                        inner_w = bw
                except Exception:
                    pass

            if inner_w < 1:
                inner_w = 1

            # Бокс обгортає текст і має padding навколо.
            box_left = text_x - pad
            box_top = text_y - pad
            box_w = inner_w + pad * 2

            # Якщо вилізли вліво/вгору, підсуваємо, і текст також.
            if box_left < 0:
                box_left = 0
                text_x = box_left + pad
            if box_top < 0:
                box_top = 0
                text_y = box_top + pad

            avail_box_w = win_w - box_left
            if avail_box_w < 1:
                avail_box_w = 1
            if box_w > avail_box_w:
                box_w = avail_box_w

            inner_w = box_w - pad * 2
            if inner_w < 1:
                inner_w = 1

            # Розкладаємо рядки під фактичну ширину.
            if wrap_chars > 0:
                lines = (txt or '').splitlines()
                if not lines:
                    lines = ['']
                fixed: List[str] = []
                for ln in lines:
                    fixed.append(_elide_line(fm, ln, inner_w))
                lines = fixed
            else:
                lines = _layout_lines_word_wrap(txt, base_font, inner_w)

            if not lines:
                lines = ['']

            # Попередній розрахунок висоти під max_lines (грубо, але ок).
            text_h = int(len(lines) * line_h)
            if text_h < 1:
                text_h = 1

            inner_h = base_h
            if text_h > inner_h:
                inner_h = text_h

            box_h = inner_h + pad * 2

            avail_box_h = win_h - box_top
            if avail_box_h < 1:
                avail_box_h = 1

            if max_box_h > 0 and box_h > max_box_h:
                box_h = max_box_h
            if box_h > avail_box_h:
                box_h = avail_box_h
            if box_h < 1:
                box_h = 1

            inner_h = box_h - pad * 2
            if inner_h < 1:
                inner_h = 1
            # Обрізаємо рядки під фактичну висоту (макс. висота або край екрану).
            max_lines = 1
            if line_h > 0:
                max_lines = int(inner_h // line_h)
                if max_lines < 1:
                    max_lines = 1

            if len(lines) > max_lines:
                kept = lines[:max_lines]
                last = kept[-1] if kept else ''
                last2 = str(last or '')

                # Додаємо трикрапку і елліпсимо під ширину.
                if not last2.endswith('…'):
                    last2 = last2.rstrip() + '…'
                last2 = _elide_line(fm, last2, inner_w)
                if kept:
                    kept[-1] = last2
                lines = kept

            txt_final = "\n".join(lines)

            # Після обрізання можемо уточнити потрібну висоту, але не перевищуючи ліміти.
            text_h2 = int(max(1, len(lines) * line_h))
            needed_inner_h = base_h
            if text_h2 > needed_inner_h:
                needed_inner_h = text_h2

            needed_box_h = needed_inner_h + pad * 2
            # не роздуваємося більше того, що вже дозволено
            if max_box_h > 0 and needed_box_h > max_box_h:
                needed_box_h = max_box_h
            if needed_box_h > avail_box_h:
                needed_box_h = avail_box_h
            if needed_box_h < 1:
                needed_box_h = 1

            box_h = int(needed_box_h)
            inner_h = int(max(1, box_h - pad * 2))

            # Якщо вилізли вниз, підсуваємо вгору (і текст також).
            if box_top + box_h > win_h:
                if box_h >= win_h:
                    box_top = 0
                    box_h = win_h
                else:
                    box_top = win_h - box_h
                text_y = box_top + pad

            # Якщо вилізли вправо, підсуваємо вліво (і текст також).
            if box_left + box_w > win_w:
                if box_w >= win_w:
                    box_left = 0
                    box_w = win_w
                else:
                    box_left = win_w - box_w
                text_x = box_left + pad

            rect = QtCore.QRect(int(box_left), int(box_top), int(box_w), int(box_h))

            bg_rgb = bg_rgb_default
            bg_a = bg_a_default
            if bool(style.use_ocr_bg_color) and it.bg_color and len(it.bg_color) == 3:
                bg_rgb = _sanitize_rgb(it.bg_color, bg_rgb_default)

            color = QtGui.QColor(int(bg_rgb[0]), int(bg_rgb[1]), int(bg_rgb[2]), int(bg_a))

            painter.setPen(QtCore.Qt.PenStyle.NoPen)
            painter.setBrush(QtGui.QBrush(color))
            painter.drawRoundedRect(rect, rr, rr)

            painter.setPen(text_pen)

            # QRectF, щоб не втрачати дробові пікселі на HiDPI.
            text_rect = QtCore.QRectF(
                float(text_x),
                float(text_y),
                float(max(1, int(inner_w))),
                float(max(1, int(inner_h))),
            )

            painter.drawText(text_rect, flags, txt_final)

        painter.end()