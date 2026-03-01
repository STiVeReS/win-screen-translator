from __future__ import annotations

import asyncio
import os
import threading
import logging
from typing import List, Optional, Tuple

from PySide6 import QtCore, QtGui, QtWidgets

from .capture import capture_region_for_ocr, capture_region_png, list_monitors, CaptureInfo
from .config import AppConfig, RoiSectionConfig, save_config
from .overlay import OverlayItem, OverlayWindow, OverlayStyle, _layout_lines_word_wrap, _wrap_text_by_chars
from .providers import ProviderManager
from .text_similarity import is_same_or_similar, normalize_text
from .text_merge import merge_close_text_regions


logger = logging.getLogger(__name__)


def _median(values: List[float]) -> float:
    vals = [float(v) for v in (values or []) if v is not None]
    if not vals:
        return 0.0
    vals.sort()
    n = int(len(vals))
    mid = int(n // 2)
    if (n % 2) == 1:
        return float(vals[mid])
    return (float(vals[mid - 1]) + float(vals[mid])) / 2.0

def _percentile(values: List[float], q: float) -> float:
    vals = [float(v) for v in (values or []) if v is not None]
    if not vals:
        return 0.0
    try:
        qq = float(q)
    except Exception:
        qq = 0.5
    if qq < 0.0:
        qq = 0.0
    if qq > 1.0:
        qq = 1.0
    vals.sort()
    if len(vals) == 1:
        return float(vals[0])
    pos = qq * float(len(vals) - 1)
    lo = int(pos)
    hi = int(min(len(vals) - 1, lo + 1))
    frac = float(pos - lo)
    return float(vals[lo] * (1.0 - frac) + vals[hi] * frac)


def _clamp_int(v: int, lo: int, hi: int) -> int:
    try:
        x = int(v)
    except Exception:
        x = int(lo)
    if x < int(lo):
        x = int(lo)
    if x > int(hi):
        x = int(hi)
    return x


def _coerce_rgb(value, default: List[int]) -> List[int]:
    if isinstance(value, (list, tuple)) and len(value) == 3:
        out: List[int] = []
        for i in range(3):
            try:
                out.append(_clamp_int(int(value[i]), 0, 255))
            except Exception:
                out.append(int(default[i]))
        return out
    return list(default)


def _rgb_to_hex(rgb: List[int]) -> str:
    try:
        r = _clamp_int(int(rgb[0]), 0, 255)
        g = _clamp_int(int(rgb[1]), 0, 255)
        b = _clamp_int(int(rgb[2]), 0, 255)
        return f"#{r:02X}{g:02X}{b:02X}"
    except Exception:
        return "#000000"


def _hex_to_rgb(text: str) -> Optional[List[int]]:
    s = str(text or '').strip()
    if not s:
        return None
    if s.startswith('#'):
        s = s[1:]
    if len(s) != 6:
        return None
    try:
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
        return [_clamp_int(r, 0, 255), _clamp_int(g, 0, 255), _clamp_int(b, 0, 255)]
    except Exception:
        return None


def _copy_style(base: OverlayStyle) -> OverlayStyle:
    st = OverlayStyle()
    st.font_family = str(getattr(base, 'font_family', 'Segoe UI') or 'Segoe UI')
    st.font_size = float(getattr(base, 'font_size', 14.0) or 14.0)
    st.padding = int(getattr(base, 'padding', 6) or 6)
    st.round_radius = int(getattr(base, 'round_radius', 8) or 8)
    st.bg_color = list(getattr(base, 'bg_color', [0, 0, 0]) or [0, 0, 0])
    st.bg_opacity = int(getattr(base, 'bg_opacity', 170) or 170)
    st.text_color = list(getattr(base, 'text_color', [255, 255, 255]) or [255, 255, 255])
    st.text_opacity = int(getattr(base, 'text_opacity', 235) or 235)
    st.use_ocr_bg_color = bool(getattr(base, 'use_ocr_bg_color', False))
    st.max_chars_per_line = int(getattr(base, 'max_chars_per_line', 0) or 0)
    st.max_box_height = int(getattr(base, 'max_box_height', 0) or 0)
    return st


def _build_effective_section_style(base: OverlayStyle, section: RoiSectionConfig) -> OverlayStyle:
    st = _copy_style(base)
    if not bool(getattr(section, 'overlay_custom', True)):
        return st

    if getattr(section, 'overlay_font_family', None) is not None:
        fam = str(section.overlay_font_family or '').strip()
        if fam:
            st.font_family = fam

    if getattr(section, 'overlay_font_size', None) is not None:
        try:
            st.font_size = float(section.overlay_font_size)
        except Exception:
            pass

    if getattr(section, 'overlay_padding', None) is not None:
        try:
            st.padding = int(section.overlay_padding)
        except Exception:
            pass

    if getattr(section, 'overlay_round_radius', None) is not None:
        try:
            st.round_radius = int(section.overlay_round_radius)
        except Exception:
            pass

    if getattr(section, 'overlay_bg_color', None) is not None:
        st.bg_color = _coerce_rgb(section.overlay_bg_color, st.bg_color)

    if getattr(section, 'overlay_bg_opacity', None) is not None:
        try:
            st.bg_opacity = _clamp_int(int(section.overlay_bg_opacity), 0, 255)
        except Exception:
            pass

    if getattr(section, 'overlay_text_color', None) is not None:
        st.text_color = _coerce_rgb(section.overlay_text_color, st.text_color)

    if getattr(section, 'overlay_text_opacity', None) is not None:
        try:
            st.text_opacity = _clamp_int(int(section.overlay_text_opacity), 0, 255)
        except Exception:
            pass

    if getattr(section, 'overlay_use_ocr_bg_color', None) is not None:
        st.use_ocr_bg_color = bool(section.overlay_use_ocr_bg_color)

    if getattr(section, 'overlay_max_chars_per_line', None) is not None:
        try:
            st.max_chars_per_line = int(section.overlay_max_chars_per_line)
        except Exception:
            pass

    if getattr(section, 'overlay_max_box_height', None) is not None:
        try:
            st.max_box_height = int(section.overlay_max_box_height)
        except Exception:
            st.max_box_height = 0

    return st


def _font_metrics_for_device(font: QtGui.QFont, device=None) -> QtGui.QFontMetricsF:
    """Повертає QFontMetricsF, максимально наближений до того, що бачить оверлей.

    На HiDPI/різних моніторах метрики без paint-device інколи дають інші числа,
    через що авто-підбір виходить "кривим".
    """
    try:
        if device is not None:
            return QtGui.QFontMetricsF(font, device)
    except Exception:
        pass
    return QtGui.QFontMetricsF(font)

class _RoiWorkerSignals(QtCore.QObject):
    finished = QtCore.Signal(object, object, object, object)  # (section_id, capture_info, items, meta)
    failed = QtCore.Signal(object, str)  # (section_id, msg)


class RoiTranslateWorker(threading.Thread):
    def __init__(
        self,
        cfg: AppConfig,
        section: RoiSectionConfig,
        signals: _RoiWorkerSignals,
        prev_signature: Optional[str] = None,
        prev_translations: Optional[List[str]] = None,
    ):
        super().__init__(daemon=True)
        self.cfg = cfg
        self.section = section
        self.signals = signals
        self.prev_signature = prev_signature
        self.prev_translations = prev_translations

    def run(self) -> None:
            sid = self.section.id
            try:
                ocr_provider = (self.cfg.ocr_provider or '').strip().lower()
                if ocr_provider == 'ocrspace':
                    image_bytes, cap = capture_region_for_ocr(
                        monitor_index=self.section.monitor_index,
                        rel_x=self.section.x,
                        rel_y=self.section.y,
                        width=self.section.width,
                        height=self.section.height,
                    )
                    cap_ext = 'jpg'
                else:
                    image_bytes, cap = capture_region_png(
                        monitor_index=self.section.monitor_index,
                        rel_x=self.section.x,
                        rel_y=self.section.y,
                        width=self.section.width,
                        height=self.section.height,
                    )
                    cap_ext = 'png'

                try:
                    if bool(getattr(self.cfg, 'debug_ocr_log', False)):
                        dbg_dir = os.path.join(self.cfg.config_dir(), 'debug')
                        os.makedirs(dbg_dir, exist_ok=True)
                        safe = (sid or 'roi').replace('/', '_').replace('\\', '_')
                        dbg_path = os.path.join(dbg_dir, f'last_capture_{safe}.{cap_ext}')
                        with open(dbg_path, 'wb') as f:
                            f.write(image_bytes)
                        logger.info('Debug capture saved: %s (%d KB)', dbg_path, int(len(image_bytes) / 1024))
                except Exception:
                    pass

                pm = ProviderManager(data_dir=self.cfg.config_dir(), ocrspace_api_key=self.cfg.ocrspace_api_key)
                pm.configure(
                    google_api_key=self.cfg.google_api_key,
                    ocr_provider=self.cfg.ocr_provider,
                    translation_provider=self.cfg.translation_provider,
                    rapidocr_models_dir=self.cfg.rapidocr_models_dir,
                    ocrspace_api_key=self.cfg.ocrspace_api_key,
                    data_dir=self.cfg.config_dir(),
                )
                pm.set_rapidocr_confidence(self.cfg.rapidocr_min_confidence)
                pm.set_rapidocr_box_thresh(self.cfg.rapidocr_box_thresh)
                pm.set_rapidocr_unclip_ratio(self.cfg.rapidocr_unclip_ratio)

                async def pipeline():
                    regions = await pm.recognize_text(image_bytes, self.cfg.source_lang)

                    # -- 1. ЗБЕРІГАЄМО ОРИГІНАЛЬНІ ВИСОТИ БОКСІВ ДО ЗЛИТТЯ --
                    raw_heights_cap = []
                    sy = float(cap.scale_y)
                    for r in regions:
                        try:
                            rect = getattr(r, 'rect', None)
                            h = 0.0
                            if isinstance(rect, dict):
                                if 'bottom' in rect and 'top' in rect:
                                    h = float(rect['bottom']) - float(rect.get('top', 0))
                                else:
                                    h = float(rect.get('height', 0))
                            elif isinstance(rect, (list, tuple)) and len(rect) == 4:
                                h = float(rect[3])
                            
                            if h > 4.0:
                                raw_heights_cap.append(h * sy)
                        except Exception:
                            pass

                    # -- 2. ЗЛІПЛЮЄМО БЛИЗЬКІ БОКСИ --
                    try:
                        if bool(getattr(self.cfg, 'merge_close_text_regions', True)):
                            before_n = len(regions)
                            regions = merge_close_text_regions(
                                regions,
                                enabled=True,
                                x_gap_ratio=float(getattr(self.cfg, 'merge_x_gap_ratio', 1.25) or 1.25),
                                line_y_ratio=float(getattr(self.cfg, 'merge_line_y_ratio', 0.70) or 0.70),
                                merge_vertical=bool(getattr(self.cfg, 'merge_vertical_lines', True)),
                            )
                            after_n = len(regions)
                            if after_n != before_n:
                                logger.info('ROI OCR merge: section=%s %d -> %d regions', self.section.name, before_n, after_n)
                    except Exception:
                        pass

                    try:
                        logger.info(
                            "ROI OCR done: section=%s provider=%s regions=%d capture=%dx%d at (%d,%d)",
                            self.section.name,
                            pm.last_ocr_provider_name or self.cfg.ocr_provider,
                            len(regions),
                            cap.width, cap.height,
                            cap.left, cap.top,
                        )
                    except Exception:
                        pass

                    texts = [r.text for r in regions]

                    def _sig(parts: List[str]) -> str:
                        cleaned: List[str] = []
                        for p in (parts or []):
                            s = (p or '').strip()
                            cleaned.append(s)
                        joined = '\n'.join(cleaned)
                        return f"{self.cfg.source_lang}->{self.cfg.target_lang}|{joined}".strip()

                    cleaned_src: List[str] = []
                    for p in (texts or []):
                        s = (p or '').strip()
                        if s:
                            cleaned_src.append(s)

                    src_text = '\n'.join(cleaned_src).strip()
                    src_norm = normalize_text(src_text)

                    signature = _sig(texts)

                    translations: List[str] = []
                    can_reuse = False
                    if bool(getattr(self.cfg, 'skip_translate_if_same_text', True)):
                        if self.prev_signature is not None and str(self.prev_signature) == str(signature):
                            if self.prev_translations is not None and len(self.prev_translations) == len(texts):
                                can_reuse = True

                    if can_reuse:
                        translations = list(self.prev_translations or [])
                        try:
                            logger.info('ROI Translate skipped (same text): section=%s items=%d', self.section.name, len(translations))
                        except Exception:
                            pass
                    else:
                        if not src_norm:
                            translations = []
                        else:
                            translations = await pm.translate_text(texts, self.cfg.source_lang, self.cfg.target_lang)

                    try:
                        logger.info(
                            "ROI Translate done: section=%s provider=%s items=%d %s->%s",
                            self.section.name,
                            pm.last_translation_provider_name or self.cfg.translation_provider,
                            len(translations),
                            self.cfg.source_lang,
                            self.cfg.target_lang,
                        )
                    except Exception:
                        pass

                    items: List[OverlayItem] = []

                    sx = float(cap.scale_x)
                    sy = float(cap.scale_y)

                    for r, t in zip(regions, translations):
                        if not t or not t.strip():
                            continue

                        rect_raw = getattr(r, 'rect', None)
                        rect = rect_raw

                        if isinstance(rect_raw, (list, tuple)) and len(rect_raw) == 4:
                            rect = {
                                'left': rect_raw[0],
                                'top': rect_raw[1],
                                'width': rect_raw[2],
                                'height': rect_raw[3],
                            }

                        if not isinstance(rect, dict):
                            continue

                        left = int(float(rect.get('left', 0)) * sx)
                        top = int(float(rect.get('top', 0)) * sy)

                        if rect.get('right') is not None and rect.get('bottom') is not None:
                            right = int(float(rect.get('right', 0)) * sx)
                            bottom = int(float(rect.get('bottom', 0)) * sy)
                        else:
                            w = float(rect.get('width', 0))
                            h = float(rect.get('height', 0))
                            right = int((float(rect.get('left', 0)) + w) * sx)
                            bottom = int((float(rect.get('top', 0)) + h) * sy)

                        if left < 0:
                            left = 0
                        if top < 0:
                            top = 0
                        if right > int(cap.width):
                            right = int(cap.width)
                        if bottom > int(cap.height):
                            bottom = int(cap.height)

                        if right <= left or bottom <= top:
                            continue

                        items.append(
                            OverlayItem(
                                left=left,
                                top=top,
                                right=right,
                                bottom=bottom,
                                text=t,
                                bg_color=r.bg_color,
                            )
                        )

                    meta = {
                        'signature': signature,
                        'translations': translations,
                        'src_text': src_text,
                        'src_norm': src_norm,
                        'raw_heights': raw_heights_cap, # Передаємо висоти у метадані
                    }

                    return items, meta

                items, meta = asyncio.run(pipeline())
                self.signals.finished.emit(sid, cap, items, meta)
            except Exception as e:
                self.signals.failed.emit(sid, str(e))

class RoiSelectionDialog(QtWidgets.QDialog):
    """Модальний селектор ROI без ручних event-loop, щоб не «заморожувати» весь UI."""

    def __init__(self, screen: QtGui.QScreen, parent: Optional[QtWidgets.QWidget] = None):
        super().__init__(parent)
        self._screen = screen
        self._origin: Optional[QtCore.QPoint] = None
        self._current: Optional[QtCore.QPoint] = None
        self._selected: Optional[QtCore.QRect] = None

        self.setWindowFlag(QtCore.Qt.WindowType.FramelessWindowHint, True)
        self.setWindowFlag(QtCore.Qt.WindowType.WindowStaysOnTopHint, True)
        self.setWindowFlag(QtCore.Qt.WindowType.Tool, True)
        self.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground, True)

        self.setCursor(QtCore.Qt.CursorShape.CrossCursor)
        self.setFocusPolicy(QtCore.Qt.FocusPolicy.StrongFocus)

        g = self._screen.geometry()
        self.setGeometry(g)

    def selected_rect(self) -> Optional[QtCore.QRect]:
        return self._selected

    def showEvent(self, event: QtGui.QShowEvent) -> None:
        super().showEvent(event)
        try:
            self.grabKeyboard()
            self.grabMouse()
        except Exception:
            pass

        try:
            self.activateWindow()
            self.raise_()
        except Exception:
            pass

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        try:
            self.releaseKeyboard()
            self.releaseMouse()
        except Exception:
            pass
        super().closeEvent(event)

    def keyPressEvent(self, event: QtGui.QKeyEvent) -> None:
        if event.key() == QtCore.Qt.Key.Key_Escape:
            self.reject()
            return

        if event.key() in (QtCore.Qt.Key.Key_Return, QtCore.Qt.Key.Key_Enter):
            if self._selected is not None:
                self.accept()
                return

        super().keyPressEvent(event)

    def mousePressEvent(self, event: QtGui.QMouseEvent) -> None:
        if event.button() == QtCore.Qt.MouseButton.RightButton:
            self.reject()
            return

        if event.button() != QtCore.Qt.MouseButton.LeftButton:
            return

        self._origin = event.position().toPoint()
        self._current = self._origin
        self._selected = None
        self.update()

    def mouseMoveEvent(self, event: QtGui.QMouseEvent) -> None:
        if self._origin is None:
            return
        self._current = event.position().toPoint()
        self.update()

    def mouseReleaseEvent(self, event: QtGui.QMouseEvent) -> None:
        if self._origin is None:
            return

        self._current = event.position().toPoint()
        rect = QtCore.QRect(self._origin, self._current).normalized()

        self._origin = None
        self._current = None

        if rect.width() < 10 or rect.height() < 10:
            self._selected = None
            self.update()
            return

        self._selected = rect
        self.update()

        # Для швидкого додавання робимо accept одразу на відпусканні.
        self.accept()

    def paintEvent(self, event: QtGui.QPaintEvent) -> None:
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)

        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        painter.setBrush(QtGui.QColor(0, 0, 0, 120))
        painter.drawRect(self.rect())

        # Підказка
        painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 230)))
        painter.setFont(QtGui.QFont('Segoe UI', 11))
        painter.drawText(20, 28, 'Виділи зону мишкою. Esc або ПКМ = скасувати.')

        if self._origin is not None and self._current is not None:
            rect = QtCore.QRect(self._origin, self._current).normalized()
        else:
            rect = self._selected

        if rect is not None:
            painter.setBrush(QtGui.QColor(0, 0, 0, 0))
            pen = QtGui.QPen(QtGui.QColor(0, 180, 255, 230))
            pen.setWidth(2)
            painter.setPen(pen)
            painter.drawRect(rect)

            txt = f"{rect.width()} x {rect.height()}"
            painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 240)))
            painter.setFont(QtGui.QFont('Segoe UI', 11))
            painter.drawText(rect.topLeft() + QtCore.QPoint(6, 18), txt)

        painter.end()


def _pick_screen_for_monitor(app: QtWidgets.QApplication, monitor_index: int) -> Tuple[QtGui.QScreen, float, dict]:
    mons = list_monitors()
    mon = None
    for m in mons:
        if int(m.get('index', 0)) == int(monitor_index):
            mon = m
            break

    screens = app.screens() or []
    screen = app.primaryScreen()

    if mon is None:
        dpr = 1.0
        if screen is not None:
            dpr = float(screen.devicePixelRatio() or 1.0)
        return screen, dpr, {'index': int(monitor_index), 'left': 0, 'top': 0, 'width': 0, 'height': 0}

    best = None
    best_score = 10 ** 18
    for s in screens:
        try:
            g = s.geometry()
            dpr_s = float(s.devicePixelRatio() or 1.0)
            phys_left = int(round(g.left() * dpr_s))
            phys_top = int(round(g.top() * dpr_s))
            phys_w = int(round(g.width() * dpr_s))
            phys_h = int(round(g.height() * dpr_s))

            score = abs(phys_left - int(mon.get('left', 0)))
            score += abs(phys_top - int(mon.get('top', 0)))
            score += abs(phys_w - int(mon.get('width', 0)))
            score += abs(phys_h - int(mon.get('height', 0)))

            if score < best_score:
                best = s
                best_score = score
        except Exception:
            continue

    if best is not None:
        screen = best

    dpr = 1.0
    if screen is not None:
        dpr = float(screen.devicePixelRatio() or 1.0)

    return screen, dpr, mon


def _clamp_interval(cfg: AppConfig, interval_ms: int) -> int:
    interval = int(interval_ms or 0)

    if interval < 300:
        interval = 300

    ocr_provider = (cfg.ocr_provider or '').strip().lower()
    if ocr_provider == 'ocrspace' and interval < 1500:
        interval = 1500

    if interval > 60000:
        interval = 60000

    return interval


class RoiSectionController(QtCore.QObject):
    def __init__(self, app: QtWidgets.QApplication, cfg: AppConfig, section: RoiSectionConfig):
        super().__init__()
        self.app = app
        self.cfg = cfg
        self.section = section

        self.section.ensure_id()

        self.overlay = OverlayWindow()
        self.overlay.hide()

        self._overlay_enabled_global = True

        self.signals = _RoiWorkerSignals()
        self.signals.finished.connect(self._on_done)
        self.signals.failed.connect(self._on_failed)

        self._busy = False
        self._active = False

        self._last_items: Optional[List[OverlayItem]] = None
        self._last_ocr_signature: Optional[str] = None
        self._last_translations: Optional[List[str]] = None

        self._base_overlay_style: Optional[OverlayStyle] = None
        self._overlay_style: Optional[OverlayStyle] = None
        self._dpr: float = 1.0
        self._inv_dpr: float = 1.0

        # Схожість тексту: якщо ≥ threshold, не оновлюємо бокс і не перекладаємо заново
        self._last_src_text: str = ""
        self._last_src_norm: str = ""
        self._similarity_threshold: float = 0.90

        self._timer = QtCore.QTimer(self)
        self._timer.setTimerType(QtCore.Qt.TimerType.PreciseTimer)
        self._timer.timeout.connect(self._tick)

        self._update_overlay_geometry()

    def set_overlay_enabled_global(self, enabled: bool) -> None:
        self._overlay_enabled_global = bool(enabled)
        if not self._overlay_enabled_global:
            self.overlay.hide()
            return

        if self.section.enabled and self._last_items:
            self.overlay.show()

    def set_overlay_style(self, style: OverlayStyle) -> None:
        # style тут приходить як «база» (глобальні налаштування).
        self._base_overlay_style = style
        eff = _build_effective_section_style(style, self.section)
        self._overlay_style = eff
        try:
            self.overlay.set_style(eff)
        except Exception:
            pass

    def _maybe_auto_init_font_and_box(self, items_dip: List[OverlayItem], raw_heights_dip: Optional[List[float]] = None) -> None:
        if not bool(getattr(self.section, 'overlay_custom', True)):
            return

        if not items_dip:
            return

        need_font = getattr(self.section, 'overlay_font_size', None) is None
        need_box = getattr(self.section, 'overlay_max_box_height', None) is None
        if not (need_font or need_box):
            return

        need_save = False

        try:
            roi_h = int(self.overlay.height() or 0)
        except Exception:
            roi_h = 0
        if roi_h <= 0:
            try:
                roi_h = int(round(float(self.section.height) / float(self._dpr or 1.0)))
            except Exception:
                roi_h = 0
        if roi_h <= 0:
            roi_h = 1

        heights: List[float] = []
        valid_text_len = 0
        
        # ВИКОРИСТОВУЄМО ОРИГІНАЛЬНІ ВИСОТИ (до злиття рядків), щоб шрифт не був велетенським
        if raw_heights_dip and len(raw_heights_dip) > 0:
            for h in raw_heights_dip:
                if h >= 6.0 and h <= float(roi_h * 1.2):
                    heights.append(h)
            for it in items_dip:
                valid_text_len += len((it.text or '').strip())
                
        # Fallback, якщо оригінальних висот немає
        if not heights:
            for it in items_dip:
                try:
                    h = float(int(it.bottom) - int(it.top))
                    w = float(int(it.right) - int(it.left))
                except Exception:
                    continue
                if h < 6 or h > float(roi_h * 1.2):
                    continue
                if w < 10:
                    continue
                heights.append(h)
                valid_text_len += len((it.text or '').strip())

        if not heights or valid_text_len < 3:
            return

        med_h = _median(heights)
        if med_h <= 0.0:
            med_h = float(min(28, roi_h))

        line_h_est = med_h

        # Оцінка типового числа рядків
        est_lines = 1
        max_box_h = 0.0
        for it in items_dip:
            try:
                hh = float(int(it.bottom) - int(it.top))
                if hh > max_box_h:
                    max_box_h = hh
            except Exception:
                pass
                
        if max_box_h > float(line_h_est) * 1.4:
            est_lines = int(round(max_box_h / float(line_h_est)))
            
        if est_lines < 1:
            est_lines = 1

        # 1) Авто розмір шрифту
        if need_font:
            try:
                target_text_h = float(line_h_est) * 0.70
                if target_text_h < 8.0:
                    target_text_h = 8.0
                if target_text_h > 72.0:
                    target_text_h = 72.0

                scale = float(getattr(self.cfg, 'font_scale', 1.0) or 1.0)
                if scale <= 0:
                    scale = 1.0

                point_size_used = float(target_text_h) / 1.3333333333
                base_size = float(point_size_used) / float(scale)

                if base_size < 8.0:
                    base_size = 8.0
                if base_size > 64.0:
                    base_size = 64.0

                self.section.overlay_font_size = float(round(base_size, 1))
                need_save = True
            except Exception:
                pass

        # 2) Авто максимальна висота боксу
        if need_box:
            try:
                base = self._base_overlay_style
                if base is None:
                    base = OverlayStyle()
                eff = _build_effective_section_style(base, self.section)

                pad = int(getattr(eff, 'padding', 6) or 6)
                calc_pad = min(pad, 6)

                lines_cap = int(est_lines) + 1
                if lines_cap < 2:
                    lines_cap = 2
                if lines_cap > 5:
                    lines_cap = 5

                picked = (float(line_h_est) * float(lines_cap)) + float(calc_pad)
                picked = picked * 1.05

                if picked < 20.0:
                    picked = 20.0
                if picked > float(roi_h):
                    picked = float(roi_h)

                self.section.overlay_max_box_height = int(round(picked))
                need_save = True
            except Exception:
                pass

        if need_save:
            try:
                save_config(self.cfg)
            except Exception:
                pass
            try:
                if self._base_overlay_style is not None:
                    eff = _build_effective_section_style(self._base_overlay_style, self.section)
                    self._overlay_style = eff
                    self.overlay.set_style(eff)
            except Exception:
                pass
    
    def _update_overlay_geometry(self) -> None:
        screen, dpr, mon = _pick_screen_for_monitor(self.app, self.section.monitor_index)
        if screen is None:
            return

        try:
            self._dpr = float(dpr or 1.0)
        except Exception:
            self._dpr = 1.0
        if self._dpr <= 0:
            self._dpr = 1.0
        self._inv_dpr = 1.0 / float(self._dpr)

        g = screen.geometry()  # DIP

        # ROI координати зберігаються відносно монітора у фізичних пікселях.
        x_dip = int(round(float(self.section.x) / float(dpr or 1.0)))
        y_dip = int(round(float(self.section.y) / float(dpr or 1.0)))
        w_dip = int(round(float(self.section.width) / float(dpr or 1.0)))
        h_dip = int(round(float(self.section.height) / float(dpr or 1.0)))

        if w_dip < 1:
            w_dip = 1
        if h_dip < 1:
            h_dip = 1

        left = int(g.left()) + int(x_dip)
        top = int(g.top()) + int(y_dip)

        self.overlay.set_overlay_geometry(left, top, w_dip, h_dip)

        try:
            logger.info(
                "ROI Overlay geometry: section=%s monitor=%s dpr=%.3f dip=(%d,%d %dx%d) phys=(%d,%d %dx%d)",
                self.section.name,
                self.section.monitor_index,
                float(dpr or 1.0),
                left, top, w_dip, h_dip,
                int(mon.get('left', 0)) + int(self.section.x),
                int(mon.get('top', 0)) + int(self.section.y),
                int(self.section.width),
                int(self.section.height),
            )
        except Exception:
            pass

    def start(self) -> None:
        if self._active:
            return
        self._active = True

        interval = _clamp_interval(self.cfg, self.section.interval_ms)
        self._timer.start(interval)

        # перший прохід одразу
        self._tick()

    def stop(self) -> None:
        self._active = False
        self._timer.stop()
        self.overlay.hide()

    def is_active(self) -> bool:
        return bool(self._active)

    def hide(self) -> None:
        self.overlay.hide()

    def update_once(self) -> None:
        self._tick()

    def apply_section(self, section: RoiSectionConfig) -> None:
        self.section = section
        self.section.ensure_id()
        self._update_overlay_geometry()

        # Переобчислити ефективний стиль, якщо є базовий.
        try:
            if self._base_overlay_style is not None:
                eff = _build_effective_section_style(self._base_overlay_style, self.section)
                self._overlay_style = eff
                self.overlay.set_style(eff)
        except Exception:
            pass

        # якщо вимкнули секцію у конфігу
        if not self.section.enabled:
            self.overlay.hide()

        # якщо активна і поміняли інтервал, оновлюємо
        if self._active:
            interval = _clamp_interval(self.cfg, self.section.interval_ms)
            self._timer.start(interval)

    def _hide_all_overlays_now(self) -> None:
        """Ховає всі OverlayWindow перед захопленням ROI, щоб OCR не бачив старі бокси."""
        try:
            for w in self.app.topLevelWidgets():
                try:
                    if isinstance(w, OverlayWindow):
                        w.hide()
                except Exception:
                    continue

            try:
                self.app.processEvents()
            except Exception:
                pass

            try:
                QtCore.QThread.msleep(35)
            except Exception:
                pass
        except Exception:
            pass

    def _restore_own_overlay_if_needed(self) -> None:
        """Повертає поточний бокс цієї секції, якщо він був і його треба залишити."""
        if not self._overlay_enabled_global:
            return
        if not self.section.enabled:
            return
        if not self._last_items:
            return

        try:
            if self._overlay_style is not None:
                self.overlay.set_style(self._overlay_style)
        except Exception:
            pass

        try:
            self.overlay.show()
            self.overlay.raise_()
        except Exception:
            pass

    def _clear_overlay(self) -> None:
        try:
            self.overlay.set_items([], font_scale=self.cfg.font_scale)
        except Exception:
            pass
        try:
            self.overlay.hide()
        except Exception:
            pass

    def _tick(self) -> None:
        if not self.section.enabled:
            return
        if self._busy:
            return

        need_hide = True
        try:
            if bool(getattr(self.cfg, 'exclude_overlay_from_capture', True)):
                if bool(self.overlay.is_excluded_from_capture()):
                    need_hide = False
        except Exception:
            need_hide = True

        if need_hide:
            self._hide_all_overlays_now()

        self._busy = True
        worker = RoiTranslateWorker(
            self.cfg,
            self.section,
            self.signals,
            prev_signature=self._last_ocr_signature,
            prev_translations=self._last_translations,
        )
        worker.start()

    @QtCore.Slot(object, object, object, object)
    def _on_done(self, section_id, cap: CaptureInfo, items: List[OverlayItem], meta):
            if section_id != self.section.id:
                return
            self._busy = False

            src_text = ""
            src_norm = ""
            raw_heights_cap = []
            try:
                if isinstance(meta, dict):
                    src_text = str(meta.get('src_text') or '')
                    src_norm = str(meta.get('src_norm') or '')
                    raw_heights_cap = meta.get('raw_heights') or []
            except Exception:
                src_text = ""
                src_norm = ""

            # 1) Якщо текст у зоні зник: прибираємо бокс і чистимо кеш
            if not src_norm:
                self._last_src_text = ""
                self._last_src_norm = ""
                self._last_items = None
                self._last_ocr_signature = None
                self._last_translations = None
                self._clear_overlay()
                return

            # 2) Якщо ≥90% схожий на попередній: НЕ перекладаємо/НЕ чіпаємо бокс
            if self._last_src_norm and is_same_or_similar(src_norm, self._last_src_norm, threshold=self._similarity_threshold):
                self._restore_own_overlay_if_needed()
                return

            # 3) Новий текст: оновлюємо кеш
            self._last_src_text = src_text
            self._last_src_norm = src_norm

            try:
                if isinstance(meta, dict):
                    self._last_ocr_signature = meta.get('signature')
                    self._last_translations = meta.get('translations')
            except Exception:
                pass

            if not items:
                self._last_items = None
                self._clear_overlay()
                return

            self._last_items = items

            if not self._overlay_enabled_global:
                return
            if not self.section.enabled:
                return

            if self._overlay_style is not None:
                try:
                    self.overlay.set_style(self._overlay_style)
                except Exception:
                    pass

            items_fixed: List[OverlayItem] = []

            inv_x = float(self._inv_dpr or 1.0)
            inv_y = float(self._inv_dpr or 1.0)
            try:
                ov_w = int(self.overlay.width() or 0)
                ov_h = int(self.overlay.height() or 0)
                cap_w = int(getattr(cap, 'width', 0) or 0)
                cap_h = int(getattr(cap, 'height', 0) or 0)

                if ov_w > 0 and cap_w > 0:
                    inv_x = float(ov_w) / float(cap_w)
                if ov_h > 0 and cap_h > 0:
                    inv_y = float(ov_h) / float(cap_h)
            except Exception:
                pass

            for it in items:
                items_fixed.append(
                    OverlayItem(
                        left=int(round(int(it.left) * inv_x)),
                        top=int(round(int(it.top) * inv_y)),
                        right=int(round(int(it.right) * inv_x)),
                        bottom=int(round(int(it.bottom) * inv_y)),
                        text=it.text,
                        bg_color=it.bg_color,
                    )
                )

            raw_heights_dip = []
            try:
                raw_heights_dip = [float(h) * inv_y for h in raw_heights_cap]
            except Exception:
                pass

            try:
                self._maybe_auto_init_font_and_box(items_fixed, raw_heights_dip)
            except Exception:
                pass

            self.overlay.set_items(items_fixed, font_scale=self.cfg.font_scale)
            self.overlay.show()
            try:
                self.overlay.raise_()
            except Exception:
                pass
    
    @QtCore.Slot(object, str)
    def _on_failed(self, section_id, msg: str):
        if section_id != self.section.id:
            return
        self._busy = False

        # Якщо впало, ми оверлеї ховали перед capture.
        # Краще повернути попередній бокс, ніж залишити пустоту.
        self._restore_own_overlay_if_needed()

        try:
            logger.error('ROI section failed: %s (%s): %s', self.section.name, self.section.id, msg)
        except Exception:
            pass


def select_roi_on_monitor(app: QtWidgets.QApplication, monitor_index: int, parent: Optional[QtWidgets.QWidget] = None) -> Optional[Tuple[int, int, int, int]]:
    """Повертає ROI (x,y,w,h) у ФІЗИЧНИХ пікселях відносно монітора."""

    screen, dpr, mon = _pick_screen_for_monitor(app, monitor_index)
    if screen is None:
        return None

    dlg = RoiSelectionDialog(screen, parent=parent)
    rc = dlg.exec()
    if rc != int(QtWidgets.QDialog.DialogCode.Accepted):
        return None

    rect = dlg.selected_rect()
    if rect is None:
        return None

    x = int(round(float(rect.left()) * float(dpr or 1.0)))
    y = int(round(float(rect.top()) * float(dpr or 1.0)))
    w = int(round(float(rect.width()) * float(dpr or 1.0)))
    h = int(round(float(rect.height()) * float(dpr or 1.0)))

    mw = int(mon.get('width', 0) or 0)
    mh = int(mon.get('height', 0) or 0)

    if x < 0:
        x = 0
    if y < 0:
        y = 0
    if mw > 0 and x + w > mw:
        w = mw - x
    if mh > 0 and y + h > mh:
        h = mh - y

    if w < 1 or h < 1:
        return None

    return x, y, w, h


class RoiSectionEditDialog(QtWidgets.QDialog):
    def __init__(self, app: QtWidgets.QApplication, cfg: AppConfig, section: Optional[RoiSectionConfig] = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Секція (ROI)')
        self.setMinimumSize(1000, 720)
        try:
            self.resize(1040, 760)
        except Exception:
            pass

        self._app = app
        self._cfg = cfg

        self._section = section
        self._roi: Optional[Tuple[int, int, int, int]] = None

        if section is not None:
            self._roi = (int(section.x), int(section.y), int(section.width), int(section.height))

        root = QtWidgets.QVBoxLayout(self)

        tabs = QtWidgets.QTabWidget(self)
        root.addWidget(tabs)

        # --- TAB: Основне ---
        tab_main = QtWidgets.QWidget(self)
        tabs.addTab(tab_main, 'Основне')
        form_main = QtWidgets.QFormLayout(tab_main)

        self.name = QtWidgets.QLineEdit()
        self.enabled = QtWidgets.QCheckBox('Увімкнено')

        self.monitor = QtWidgets.QComboBox()
        for m in list_monitors():
            label = f"Монітор {m['index']}: {m['width']}x{m['height']} @ ({m['left']},{m['top']})"
            self.monitor.addItem(label, int(m['index']))

        self.interval = QtWidgets.QSpinBox()
        self.interval.setRange(300, 60000)
        self.interval.setSingleStep(200)

        self.roi_label = QtWidgets.QLabel('ROI: не вибрано')
        self.pick_btn = QtWidgets.QPushButton('Вибрати ROI на екрані…')
        self.pick_btn.clicked.connect(self._pick_roi)

        if section is None:
            self.name.setText('Секція')
            self.enabled.setChecked(True)
            self.interval.setValue(int(getattr(cfg, 'continuous_interval_ms', 2000)))

            mon_default = int(getattr(cfg, 'capture_monitor', 1) or 1)
            if mon_default <= 0:
                mon_default = 1
            idx = self.monitor.findData(mon_default)
            if idx >= 0:
                self.monitor.setCurrentIndex(idx)
        else:
            self.name.setText(section.name)
            self.enabled.setChecked(bool(section.enabled))
            self.interval.setValue(int(section.interval_ms or 2000))

            idx = self.monitor.findData(int(section.monitor_index or 1))
            if idx >= 0:
                self.monitor.setCurrentIndex(idx)

        self._sync_roi_label()

        form_main.addRow('Назва:', self.name)
        form_main.addRow('Монітор:', self.monitor)
        form_main.addRow('Оновлення (мс):', self.interval)
        form_main.addRow('Стан:', self.enabled)
        form_main.addRow(self.roi_label)
        form_main.addRow(self.pick_btn)

        # --- TAB: Оверлей ---
        tab_overlay = QtWidgets.QWidget(self)
        tabs.addTab(tab_overlay, 'Оверлей')
        form_ov = None

        self.overlay_custom = QtWidgets.QCheckBox('Власні налаштування оверлею для цієї секції')

        self.font_family = QtWidgets.QFontComboBox()
        self.font_size = QtWidgets.QDoubleSpinBox()
        self.font_size.setRange(6.0, 64.0)
        self.font_size.setSingleStep(0.5)
        self.font_size_auto = QtWidgets.QCheckBox('Авто (по тексту)')

        self.padding = QtWidgets.QSpinBox()
        self.padding.setRange(0, 200)

        self.round_radius = QtWidgets.QSpinBox()
        self.round_radius.setRange(0, 200)

        self.bg_color = QtWidgets.QLineEdit()
        self.bg_opacity = QtWidgets.QSpinBox()
        self.bg_opacity.setRange(0, 255)

        self.text_color = QtWidgets.QLineEdit()
        self.text_opacity = QtWidgets.QSpinBox()
        self.text_opacity.setRange(0, 255)

        self.use_ocr_bg = QtWidgets.QCheckBox('Брати фон з OCR (якщо доступно)')

        self.max_chars = QtWidgets.QSpinBox()
        self.max_chars.setRange(0, 500)

        self.max_box_h = QtWidgets.QSpinBox()
        self.max_box_h.setRange(0, 4000)
        self.max_box_h_auto = QtWidgets.QCheckBox('Авто (по тексту)')

        # seed values: секція → глобальні
        sec_obj: Optional[RoiSectionConfig] = section
        if sec_obj is None:
            sec_obj = RoiSectionConfig()
            sec_obj.ensure_id()
            sec_obj.monitor_index = int(self.monitor.currentData() or 1)

        # overlay_custom
        self.overlay_custom.setChecked(bool(getattr(sec_obj, 'overlay_custom', True)))

        # font family
        fam = getattr(sec_obj, 'overlay_font_family', None)
        if fam is None or str(fam).strip() == '':
            fam = str(getattr(cfg, 'overlay_font_family', 'Segoe UI') or 'Segoe UI')
        try:
            self.font_family.setCurrentFont(QtGui.QFont(str(fam)))
        except Exception:
            pass

        # font size (auto when None)
        sec_fs = getattr(sec_obj, 'overlay_font_size', None)
        if sec_fs is None:
            self.font_size_auto.setChecked(True)
            self.font_size.setValue(float(getattr(cfg, 'overlay_font_size', 14.0) or 14.0))
        else:
            self.font_size_auto.setChecked(False)
            try:
                self.font_size.setValue(float(sec_fs))
            except Exception:
                self.font_size.setValue(float(getattr(cfg, 'overlay_font_size', 14.0) or 14.0))

        # padding
        pad = getattr(sec_obj, 'overlay_padding', None)
        if pad is None:
            pad = int(getattr(cfg, 'overlay_padding', 6) or 6)
        try:
            self.padding.setValue(int(pad))
        except Exception:
            self.padding.setValue(int(getattr(cfg, 'overlay_padding', 6) or 6))

        # round radius
        rr = getattr(sec_obj, 'overlay_round_radius', None)
        if rr is None:
            rr = int(getattr(cfg, 'overlay_round_radius', 8) or 8)
        try:
            self.round_radius.setValue(int(rr))
        except Exception:
            self.round_radius.setValue(int(getattr(cfg, 'overlay_round_radius', 8) or 8))

        # colors/opacities
        bgc = getattr(sec_obj, 'overlay_bg_color', None)
        if bgc is None:
            bgc = list(getattr(cfg, 'overlay_bg_color', [0, 0, 0]) or [0, 0, 0])
        self.bg_color.setText(_rgb_to_hex(_coerce_rgb(bgc, [0, 0, 0])))

        bga = getattr(sec_obj, 'overlay_bg_opacity', None)
        if bga is None:
            bga = int(getattr(cfg, 'overlay_bg_opacity', 170) or 170)
        self.bg_opacity.setValue(int(bga))

        txc = getattr(sec_obj, 'overlay_text_color', None)
        if txc is None:
            txc = list(getattr(cfg, 'overlay_text_color', [255, 255, 255]) or [255, 255, 255])
        self.text_color.setText(_rgb_to_hex(_coerce_rgb(txc, [255, 255, 255])))

        txa = getattr(sec_obj, 'overlay_text_opacity', None)
        if txa is None:
            txa = int(getattr(cfg, 'overlay_text_opacity', 235) or 235)
        self.text_opacity.setValue(int(txa))

        uob = getattr(sec_obj, 'overlay_use_ocr_bg_color', None)
        if uob is None:
            uob = bool(getattr(cfg, 'overlay_use_ocr_bg_color', False))
        self.use_ocr_bg.setChecked(bool(uob))

        mch = getattr(sec_obj, 'overlay_max_chars_per_line', None)
        if mch is None:
            mch = int(getattr(cfg, 'overlay_max_chars_per_line', 0) or 0)
        self.max_chars.setValue(int(mch))

        mbh = getattr(sec_obj, 'overlay_max_box_height', None)
        if mbh is None:
            self.max_box_h_auto.setChecked(True)
            self.max_box_h.setValue(int(getattr(cfg, 'overlay_max_box_height', 0) or 0))
        else:
            self.max_box_h_auto.setChecked(False)
            try:
                self.max_box_h.setValue(int(mbh))
            except Exception:
                self.max_box_h.setValue(int(getattr(cfg, 'overlay_max_box_height', 0) or 0))

        # layout overlay tab (категорії у grid)
        ov_grid = QtWidgets.QGridLayout(tab_overlay)
        ov_grid.setHorizontalSpacing(12)
        ov_grid.setVerticalSpacing(12)
        ov_grid.setColumnStretch(0, 1)
        ov_grid.setColumnStretch(1, 1)

        gb_typo = QtWidgets.QGroupBox('Типографіка', tab_overlay)
        f_typo = QtWidgets.QFormLayout(gb_typo)

        gb_colors = QtWidgets.QGroupBox('Кольори', tab_overlay)
        f_col = QtWidgets.QFormLayout(gb_colors)

        gb_wrap = QtWidgets.QGroupBox('Перенос і висота', tab_overlay)
        f_wrap = QtWidgets.QFormLayout(gb_wrap)

        # Typo
        f_typo.addRow(self.overlay_custom)

        row_fs = QtWidgets.QHBoxLayout()
        row_fs.addWidget(self.font_size)
        row_fs.addWidget(self.font_size_auto)

        f_typo.addRow('Шрифт:', self.font_family)
        f_typo.addRow('Розмір шрифту:', row_fs)
        f_typo.addRow('Padding:', self.padding)
        f_typo.addRow('Round radius:', self.round_radius)
        f_typo.addRow('Фон з OCR:', self.use_ocr_bg)

        # Colors
        f_col.addRow('Колір фону (hex):', self.bg_color)
        f_col.addRow('Прозорість фону (0-255):', self.bg_opacity)
        f_col.addRow('Колір тексту (hex):', self.text_color)
        f_col.addRow('Прозорість тексту (0-255):', self.text_opacity)

        # Wrap/height
        f_wrap.addRow('Wrap (символів/рядок, 0=word wrap):', self.max_chars)

        row_mbh = QtWidgets.QHBoxLayout()
        row_mbh.addWidget(self.max_box_h)
        row_mbh.addWidget(self.max_box_h_auto)
        f_wrap.addRow('Макс. висота боксу (0=без ліміту):', row_mbh)

        ov_grid.addWidget(gb_typo, 0, 0)
        ov_grid.addWidget(gb_colors, 0, 1)
        ov_grid.addWidget(gb_wrap, 1, 0, 1, 2)
        # enable/disable logic
        self.overlay_custom.toggled.connect(self._sync_overlay_enabled)
        self.font_size_auto.toggled.connect(self._sync_overlay_enabled)
        self.max_box_h_auto.toggled.connect(self._sync_overlay_enabled)
        self._sync_overlay_enabled()

        # buttons
        btns = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Save | QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        root.addWidget(btns)

    def _sync_overlay_enabled(self) -> None:
        custom = bool(self.overlay_custom.isChecked())

        # tab overlay controls
        self.font_family.setEnabled(custom)
        self.padding.setEnabled(custom)
        self.round_radius.setEnabled(custom)
        self.bg_color.setEnabled(custom)
        self.bg_opacity.setEnabled(custom)
        self.text_color.setEnabled(custom)
        self.text_opacity.setEnabled(custom)
        self.use_ocr_bg.setEnabled(custom)
        self.max_chars.setEnabled(custom)

        if custom:
            auto_fs = bool(self.font_size_auto.isChecked())
            self.font_size.setEnabled(not auto_fs)
            self.font_size_auto.setEnabled(True)

            auto_mbh = bool(self.max_box_h_auto.isChecked())
            self.max_box_h.setEnabled(not auto_mbh)
            self.max_box_h_auto.setEnabled(True)
        else:
            self.font_size.setEnabled(False)
            self.font_size_auto.setEnabled(False)
            self.max_box_h.setEnabled(False)
            self.max_box_h_auto.setEnabled(False)

    def _sync_roi_label(self) -> None:
        if self._roi is None:
            self.roi_label.setText('ROI: не вибрано')
            return
        x, y, w, h = self._roi
        self.roi_label.setText(f'ROI: x={x}, y={y}, w={w}, h={h}')

    def _pick_roi(self) -> None:
        mon = self.monitor.currentData()
        if mon is None:
            mon = 1
        try:
            mon_idx = int(mon)
        except Exception:
            mon_idx = 1

        roi = select_roi_on_monitor(self._app, mon_idx, parent=self)
        if roi is None:
            return
        self._roi = roi
        self._sync_roi_label()

    def build_section(self) -> Optional[RoiSectionConfig]:
        if self._roi is None:
            return None

        name = (self.name.text() or '').strip()
        if not name:
            name = 'Секція'

        mon = self.monitor.currentData()
        if mon is None:
            mon = 1

        try:
            mon_idx = int(mon)
        except Exception:
            mon_idx = 1

        x, y, w, h = self._roi

        out = RoiSectionConfig()
        if self._section is not None:
            out = RoiSectionConfig.from_dict(self._section.__dict__)

        out.name = name
        out.enabled = bool(self.enabled.isChecked())
        out.monitor_index = int(mon_idx)
        out.interval_ms = int(self.interval.value())
        out.x = int(x)
        out.y = int(y)
        out.width = int(w)
        out.height = int(h)

        # Пер-секційні налаштування оверлею
        out.overlay_custom = bool(self.overlay_custom.isChecked())
        if out.overlay_custom:
            try:
                out.overlay_font_family = str(self.font_family.currentFont().family() or '').strip()
            except Exception:
                out.overlay_font_family = str(getattr(self._cfg, 'overlay_font_family', 'Segoe UI') or 'Segoe UI')

            if bool(self.font_size_auto.isChecked()):
                out.overlay_font_size = None
            else:
                try:
                    out.overlay_font_size = float(self.font_size.value())
                except Exception:
                    out.overlay_font_size = float(getattr(self._cfg, 'overlay_font_size', 14.0) or 14.0)

            out.overlay_padding = int(self.padding.value())
            out.overlay_round_radius = int(self.round_radius.value())

            bg = _hex_to_rgb(self.bg_color.text())
            if bg is None:
                bg = list(getattr(self._cfg, 'overlay_bg_color', [0, 0, 0]) or [0, 0, 0])
            out.overlay_bg_color = _coerce_rgb(bg, [0, 0, 0])
            out.overlay_bg_opacity = _clamp_int(int(self.bg_opacity.value()), 0, 255)

            tx = _hex_to_rgb(self.text_color.text())
            if tx is None:
                tx = list(getattr(self._cfg, 'overlay_text_color', [255, 255, 255]) or [255, 255, 255])
            out.overlay_text_color = _coerce_rgb(tx, [255, 255, 255])
            out.overlay_text_opacity = _clamp_int(int(self.text_opacity.value()), 0, 255)

            out.overlay_use_ocr_bg_color = bool(self.use_ocr_bg.isChecked())
            out.overlay_max_chars_per_line = int(self.max_chars.value())

            if bool(self.max_box_h_auto.isChecked()):
                out.overlay_max_box_height = None
            else:
                out.overlay_max_box_height = int(self.max_box_h.value())

        out.ensure_id()
        return out


class RoiSectionsDialog(QtWidgets.QDialog):
    def __init__(self, app: QtWidgets.QApplication, cfg: AppConfig, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Win Screen Translator – Секції (ROI)')
        self.setMinimumSize(900, 600)
        try:
            self.resize(980, 640)
        except Exception:
            pass

        self._app = app
        self._cfg = cfg

        self._table = QtWidgets.QTableWidget(self)
        self._table.setColumnCount(4)
        self._table.setHorizontalHeaderLabels(['Увімкнено', 'Назва', 'Монітор', 'Інтервал (мс)'])
        self._table.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.Stretch)
        self._table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self._table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self._reloading = False
        self._table.itemChanged.connect(self._on_item_changed)

        btn_add = QtWidgets.QPushButton('Додати')
        btn_edit = QtWidgets.QPushButton('Редагувати')
        btn_del = QtWidgets.QPushButton('Видалити')

        btn_add.clicked.connect(self._add)
        btn_edit.clicked.connect(self._edit)
        btn_del.clicked.connect(self._delete)

        btns = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Close
        )
        btns.rejected.connect(self._close_commit)

        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(self._table)

        row = QtWidgets.QHBoxLayout()
        row.addWidget(btn_add)
        row.addWidget(btn_edit)
        row.addWidget(btn_del)
        row.addStretch(1)
        layout.addLayout(row)
        layout.addWidget(btns)

        self._reload_table()

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:
        try:
            self._apply_enabled_states_from_table()
        except Exception:
            pass

        # Закриття по "X" має поводитись як "Close" (тобто застосувати зміни),
        # інакше користувач отримує "нічого не працює", а винен знову я.
        try:
            self.accept()
        except Exception:
            pass

        try:
            event.accept()
        except Exception:
            pass

    def _close_commit(self) -> None:
        try:
            self._apply_enabled_states_from_table()
        except Exception:
            pass
        self.accept()

    def _reload_table(self) -> None:
        self._reloading = True
        sections = list(self._cfg.roi_sections or [])
        self._table.setRowCount(len(sections))

        r = 0
        for s in sections:
            if isinstance(s, dict):
                s = RoiSectionConfig.from_dict(s)

            it_enabled = QtWidgets.QTableWidgetItem('')
            it_enabled.setFlags(it_enabled.flags() | QtCore.Qt.ItemFlag.ItemIsUserCheckable)
            if bool(s.enabled):
                it_enabled.setCheckState(QtCore.Qt.CheckState.Checked)
            else:
                it_enabled.setCheckState(QtCore.Qt.CheckState.Unchecked)

            it_name = QtWidgets.QTableWidgetItem(str(s.name))
            it_name.setFlags(it_name.flags() & ~QtCore.Qt.ItemFlag.ItemIsEditable)
            it_mon = QtWidgets.QTableWidgetItem(str(int(s.monitor_index or 1)))
            it_mon.setFlags(it_mon.flags() & ~QtCore.Qt.ItemFlag.ItemIsEditable)
            it_int = QtWidgets.QTableWidgetItem(str(int(s.interval_ms or 2000)))
            it_int.setFlags(it_int.flags() & ~QtCore.Qt.ItemFlag.ItemIsEditable)

            self._table.setItem(r, 0, it_enabled)
            self._table.setItem(r, 1, it_name)
            self._table.setItem(r, 2, it_mon)
            self._table.setItem(r, 3, it_int)

            # зберігаємо id як data
            it_name.setData(QtCore.Qt.ItemDataRole.UserRole, str(s.id))

            r += 1

        self._table.resizeColumnsToContents()
        self._reloading = False

    def _on_item_changed(self, item: QtWidgets.QTableWidgetItem) -> None:
        if getattr(self, '_reloading', False):
            return
        if item is None:
            return
        if int(item.column()) != 0:
            return

        try:
            self._apply_enabled_states_from_table()
            save_config(self._cfg)
        except Exception:
            pass

    def _selected_section_id(self) -> Optional[str]:
        row = self._table.currentRow()
        if row < 0:
            return None
        it = self._table.item(row, 1)
        if it is None:
            return None
        sid = it.data(QtCore.Qt.ItemDataRole.UserRole)
        if not sid:
            return None
        return str(sid)

    def _find_section(self, sid: str) -> Optional[RoiSectionConfig]:
        for s in (self._cfg.roi_sections or []):
            if isinstance(s, dict):
                s = RoiSectionConfig.from_dict(s)
            if str(s.id) == str(sid):
                return s
        return None

    def _apply_enabled_states_from_table(self) -> None:
        sections_new: List[RoiSectionConfig] = []
        rows = self._table.rowCount()
        for r in range(rows):
            it_name = self._table.item(r, 1)
            it_enabled = self._table.item(r, 0)
            if it_name is None or it_enabled is None:
                continue

            sid = it_name.data(QtCore.Qt.ItemDataRole.UserRole)
            if not sid:
                continue

            sec = self._find_section(str(sid))
            if sec is None:
                continue

            if it_enabled.checkState() == QtCore.Qt.CheckState.Checked:
                sec.enabled = True
            else:
                sec.enabled = False

            sections_new.append(sec)

        self._cfg.roi_sections = sections_new

    def _add(self) -> None:
        self._apply_enabled_states_from_table()

        dlg = RoiSectionEditDialog(self._app, self._cfg, section=None, parent=self)
        if dlg.exec() != QtWidgets.QDialog.DialogCode.Accepted:
            return

        sec = dlg.build_section()
        if sec is None:
            return

        self._cfg.roi_sections.append(sec)
        self._reload_table()

    def _edit(self) -> None:
        self._apply_enabled_states_from_table()

        sid = self._selected_section_id()
        if sid is None:
            return

        sec = self._find_section(sid)
        if sec is None:
            return

        dlg = RoiSectionEditDialog(self._app, self._cfg, section=sec, parent=self)
        if dlg.exec() != QtWidgets.QDialog.DialogCode.Accepted:
            return

        sec2 = dlg.build_section()
        if sec2 is None:
            return

        out: List[RoiSectionConfig] = []
        for s in (self._cfg.roi_sections or []):
            if isinstance(s, dict):
                s = RoiSectionConfig.from_dict(s)
            if str(s.id) == str(sec2.id):
                out.append(sec2)
            else:
                out.append(s)
        self._cfg.roi_sections = out

        self._reload_table()

    def _delete(self) -> None:
        self._apply_enabled_states_from_table()

        sid = self._selected_section_id()
        if sid is None:
            return

        out: List[RoiSectionConfig] = []
        for s in (self._cfg.roi_sections or []):
            if isinstance(s, dict):
                s = RoiSectionConfig.from_dict(s)
            if str(s.id) == str(sid):
                continue
            out.append(s)
        self._cfg.roi_sections = out

        self._reload_table()