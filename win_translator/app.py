from __future__ import annotations

import asyncio
import os
import threading
import logging
from dataclasses import dataclass
from typing import List, Optional

from PySide6 import QtCore, QtGui, QtWidgets

from .capture import (
    capture_screen_for_ocr,
    capture_screen_png,
    capture_target_window_for_ocr,
    capture_target_window_png,
    list_monitors,
    CaptureInfo,
)
from .config import AppConfig, RoiSectionConfig, load_config, save_config
from .hotkey import GlobalHotkey, HotkeySpec
from .overlay import OverlayItem, OverlayWindow, OverlayStyle
from .providers import ProviderManager
from .roi_sections import RoiSectionsDialog, RoiSectionController, select_roi_on_monitor
from .text_merge import merge_close_text_regions
from .text_fix import fix_text_batch
from .win32_window import get_window_under_cursor, get_window_title, is_window_valid

logger = logging.getLogger(__name__)


def _rgb_to_hex(rgb: object, fallback: str) -> str:
    try:
        if isinstance(rgb, (list, tuple)) and len(rgb) == 3:
            r = int(rgb[0])
            g = int(rgb[1])
            b = int(rgb[2])
            if r < 0:
                r = 0
            if r > 255:
                r = 255
            if g < 0:
                g = 0
            if g > 255:
                g = 255
            if b < 0:
                b = 0
            if b > 255:
                b = 255
            return f"#{r:02X}{g:02X}{b:02X}"
    except Exception:
        pass
    return fallback


def _hex_to_rgb(value: str, fallback: List[int]) -> List[int]:
    s = (value or '').strip()
    if not s:
        return list(fallback)

    if s.startswith('#'):
        s = s[1:]

    if len(s) != 6:
        return list(fallback)

    try:
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
        return [r, g, b]
    except Exception:
        return list(fallback)


def _coerce_rgb(value: object, fallback: List[int]) -> List[int]:
    if isinstance(value, str):
        return _hex_to_rgb(value, fallback)

    if isinstance(value, (list, tuple)) and len(value) == 3:
        try:
            r = int(value[0])
            g = int(value[1])
            b = int(value[2])
            if r < 0:
                r = 0
            if r > 255:
                r = 255
            if g < 0:
                g = 0
            if g > 255:
                g = 255
            if b < 0:
                b = 0
            if b > 255:
                b = 255
            return [r, g, b]
        except Exception:
            return list(fallback)

    return list(fallback)


class _WorkerSignals(QtCore.QObject):
    finished = QtCore.Signal(object, object, object)  # (capture_info, items, meta)
    failed = QtCore.Signal(str)


class TranslateWorker(threading.Thread):
    def __init__(
        self,
        cfg: AppConfig,
        signals: _WorkerSignals,
        target_hwnd: Optional[int] = None,
        prev_signature: Optional[str] = None,
        prev_translations: Optional[List[str]] = None,
    ):
        super().__init__(daemon=True)
        self.cfg = cfg
        self.signals = signals
        self.target_hwnd = target_hwnd
        self.prev_signature = prev_signature
        self.prev_translations = prev_translations

    def run(self) -> None:
        try:
            use_target = False
            hwnd = None
            if self.target_hwnd is not None and is_window_valid(self.target_hwnd):
                use_target = True
                hwnd = int(self.target_hwnd)

            # OCR.space має ліміт по розміру, тому там залишаємо JPEG+компресію.
            # Для локального RapidOCR (і Google Cloud) краще PNG, щоб не вбивати дрібний текст.
            if (self.cfg.ocr_provider or '').strip().lower() == 'ocrspace':
                if use_target and hwnd is not None:
                    image_bytes, cap = capture_target_window_for_ocr(hwnd)
                else:
                    image_bytes, cap = capture_screen_for_ocr(self.cfg.capture_monitor)
                _cap_ext = 'jpg'
            else:
                if use_target and hwnd is not None:
                    image_bytes, cap = capture_target_window_png(hwnd)
                else:
                    image_bytes, cap = capture_screen_png(self.cfg.capture_monitor)
                _cap_ext = 'png'

            # Дебаг: зберігаємо останній захоплений кадр, щоб ти міг перевірити,
            # що захоплюється правильний монітор і що кадр не чорний.
            try:
                if bool(getattr(self.cfg, 'debug_ocr_log', False)):
                    dbg_dir = os.path.join(self.cfg.config_dir(), 'debug')
                    os.makedirs(dbg_dir, exist_ok=True)
                    dbg_path = os.path.join(dbg_dir, f'last_capture.{_cap_ext}')
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

                # Зліплюємо близькі бокси в один рядок, інакше переклад по шматках дає кашу.
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
                            logger.info('OCR merge: %d -> %d regions', before_n, after_n)
                except Exception:
                    pass

                # Логи OCR у людському форматі (пишемо в файл logs/win-screen-translator.log)
                try:
                    logger.info(
                        "OCR done: provider=%s regions=%d capture=%dx%d at (%d,%d) scale=(%.3f,%.3f)",
                        pm.last_ocr_provider_name or self.cfg.ocr_provider,
                        len(regions),
                        cap.width, cap.height,
                        cap.left, cap.top,
                        cap.scale_x, cap.scale_y,
                    )
                    if self.cfg.debug_ocr_log and regions:
                        for i, r in enumerate(regions[:80]):
                            rect_raw = getattr(r, 'rect', None)
                            rect = rect_raw
                            if isinstance(rect_raw, (list, tuple)) and len(rect_raw) == 4:
                                rect = {
                                    'left': rect_raw[0],
                                    'top': rect_raw[1],
                                    'right': float(rect_raw[0]) + float(rect_raw[2]),
                                    'bottom': float(rect_raw[1]) + float(rect_raw[3]),
                                }
                            if not isinstance(rect, dict):
                                rect = {}
                            txt = (r.text or "").replace("\n", " ").strip()
                            if len(txt) > 160:
                                txt = txt[:160] + "…"
                            logger.info(
                                "OCR[%02d] conf=%.3f box=(%s,%s,%s,%s) text=%s",
                                i,
                                float(getattr(r, 'confidence', 0.0) or 0.0),
                                rect.get('left'), rect.get('top'), rect.get('right'), rect.get('bottom'),
                                txt,
                            )
                        if len(regions) > 80:
                            logger.info("OCR: +%d regions truncated", len(regions) - 80)

                        # І ще одним блоком, щоб було зручно читати без координат
                        try:
                            all_text = "\n".join([
                                (x.text or "").strip()
                                for x in regions
                                if (x.text or "").strip()
                            ])
                            if all_text:
                                logger.info("[OCR text]\n%s", all_text)
                        except Exception:
                            pass
                except Exception:
                    pass

                texts = [r.text for r in regions]

                # Spellcheck / виправлення OCR перед перекладом (експериментально)
                try:
                    if bool(getattr(self.cfg, 'spellcheck_enabled', False)):
                        lang = str(getattr(self.cfg, 'spellcheck_lang', '') or '').strip()
                        if not lang:
                            lang = str(self.cfg.source_lang or '').strip()

                        fixed = fix_text_batch(list(texts), lang=lang)

                        changed = 0
                        for i in range(min(len(regions), len(fixed))):
                            a = str(texts[i] or '')
                            b = str(fixed[i] or '')
                            if a != b:
                                changed += 1
                            try:
                                regions[i].text = b
                            except Exception:
                                pass

                        if changed > 0:
                            logger.info('Spellcheck: changed %d/%d regions (lang=%s)', changed, len(texts), lang)

                        texts = [r.text for r in regions]
                except Exception:
                    pass

                def _sig(parts: List[str]) -> str:
                    cleaned: List[str] = []
                    for p in (parts or []):
                        s = (p or '').strip()
                        cleaned.append(s)
                    joined = '\n'.join(cleaned)
                    return f"{self.cfg.source_lang}->{self.cfg.target_lang}|{joined}".strip()

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
                        logger.info('Translate skipped (same text): items=%d %s->%s', len(translations), self.cfg.source_lang, self.cfg.target_lang)
                    except Exception:
                        pass
                else:
                    translations = await pm.translate_text(texts, self.cfg.source_lang, self.cfg.target_lang)

                try:
                    logger.info(
                        "Translate done: provider=%s items=%d %s->%s",
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

                    # Підтримка різних форматів rect:
                    #  - dict: left/top/right/bottom або left/top/width/height
                    #  - tuple/list: (x, y, w, h)
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

                    # Деякі провайдери повертають right/bottom, деякі width/height.
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

                # Дамп overlay-елементів
                try:
                    if self.cfg.debug_ocr_log:
                        for i, it in enumerate(items[:80]):
                            txt = (it.text or "").replace("\n", " ")
                            if len(txt) > 160:
                                txt = txt[:160] + "…"
                            logger.info(
                                "OVR[%02d] box=(%d,%d,%d,%d) text=%s",
                                i, it.left, it.top, it.right, it.bottom, txt
                            )
                        if len(items) > 80:
                            logger.info("OVR: +%d items truncated", len(items) - 80)
                except Exception:
                    pass

                meta = {
                    'signature': signature,
                    'translations': translations,
                }

                return items, meta

            items, meta = asyncio.run(pipeline())
            self.signals.finished.emit(cap, items, meta)
        except Exception as e:
            self.signals.failed.emit(str(e))



class HotkeyEditor(QtWidgets.QWidget):
    """Простий редактор глобальної гарячої клавіші (Ctrl/Shift/Alt + кнопка).

    Ми зберігаємо саме Win-VK коди, бо RegisterHotKey так і живе.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        self.mod_ctrl = QtWidgets.QCheckBox('Ctrl')
        self.mod_shift = QtWidgets.QCheckBox('Shift')
        self.mod_alt = QtWidgets.QCheckBox('Alt')

        self.key = QtWidgets.QComboBox()
        self._fill_keys()

        row = QtWidgets.QHBoxLayout(self)
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(self.mod_ctrl)
        row.addWidget(self.mod_shift)
        row.addWidget(self.mod_alt)
        row.addWidget(self.key, 1)

    def _fill_keys(self) -> None:
        self.key.clear()
        self.key.addItem('Вимкнено', 0)

        # A-Z
        for code in range(0x41, 0x5B):
            self.key.addItem(chr(code), int(code))

        # 0-9
        for code in range(0x30, 0x3A):
            self.key.addItem(chr(code), int(code))

        # F1-F12
        for i in range(1, 13):
            self.key.addItem(f'F{i}', int(0x70 + (i - 1)))

    def set_value(self, ctrl: bool, shift: bool, alt: bool, vk: int) -> None:
        self.mod_ctrl.setChecked(bool(ctrl))
        self.mod_shift.setChecked(bool(shift))
        self.mod_alt.setChecked(bool(alt))

        try:
            vk_i = int(vk or 0)
        except Exception:
            vk_i = 0

        idx = self.key.findData(vk_i)
        if idx < 0:
            # Якщо у конфігу якась екзотика, не ламаємось.
            self.key.addItem(f'VK 0x{vk_i:02X}', vk_i)
            idx = self.key.findData(vk_i)
        if idx >= 0:
            self.key.setCurrentIndex(idx)

    def get_value(self) -> tuple[bool, bool, bool, int]:
        ctrl = bool(self.mod_ctrl.isChecked())
        shift = bool(self.mod_shift.isChecked())
        alt = bool(self.mod_alt.isChecked())
        vk = self.key.currentData()
        if vk is None:
            vk = 0
        try:
            vk = int(vk)
        except Exception:
            vk = 0
        return ctrl, shift, alt, vk


class SettingsDialog(QtWidgets.QDialog):
    def __init__(self, cfg: AppConfig, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Win Screen Translator – Налаштування')
        self.setMinimumSize(1100, 760)

        self.cfg = cfg

        # --- widgets ---
        self.source_lang = QtWidgets.QLineEdit(cfg.source_lang)
        self.target_lang = QtWidgets.QLineEdit(cfg.target_lang)

        self.monitor = QtWidgets.QComboBox()
        self.monitor.addItem('Усі монітори (virtual screen)', 0)
        for m in list_monitors():
            label = f"Монітор {m['index']}: {m['width']}x{m['height']} @ ({m['left']},{m['top']})"
            self.monitor.addItem(label, int(m['index']))
        cur_idx = self.monitor.findData(int(cfg.capture_monitor))
        if cur_idx >= 0:
            self.monitor.setCurrentIndex(cur_idx)

        self.ocr_provider = QtWidgets.QComboBox()
        self.ocr_provider.addItems(['ocrspace', 'googlecloud', 'rapidocr'])
        self.ocr_provider.setCurrentText(cfg.ocr_provider)

        self.translation_provider = QtWidgets.QComboBox()
        self.translation_provider.addItems(['freegoogle', 'googlecloud', 'argos'])
        self.translation_provider.setCurrentText(cfg.translation_provider)

        self.google_api_key = QtWidgets.QLineEdit(cfg.google_api_key)
        self.google_api_key.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)

        self.ocrspace_api_key = QtWidgets.QLineEdit(cfg.ocrspace_api_key)
        self.ocrspace_api_key.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)

        self.skip_translate_if_same_text = QtWidgets.QCheckBox('Не перекладати повторно, якщо текст не змінився')
        self.skip_translate_if_same_text.setChecked(bool(getattr(cfg, 'skip_translate_if_same_text', True)))

        # OCR merge
        self.merge_close_text_regions = QtWidgets.QCheckBox('Зліплювати близькі OCR-бокси (щоб рядки не ламались на шматки)')
        self.merge_close_text_regions.setChecked(bool(getattr(cfg, 'merge_close_text_regions', True)))

        self.merge_vertical_lines = QtWidgets.QCheckBox('Зліплювати сусідні рядки в один бокс (мультилайн)')
        self.merge_vertical_lines.setChecked(bool(getattr(cfg, 'merge_vertical_lines', True)))

        # Spellcheck
        self.spellcheck_enabled = QtWidgets.QCheckBox('Виправляти OCR-текст перед перекладом (spellcheck, експериментально)')
        self.spellcheck_enabled.setChecked(bool(getattr(cfg, 'spellcheck_enabled', False)))

        self.spellcheck_lang = QtWidgets.QLineEdit(str(getattr(cfg, 'spellcheck_lang', '') or '').strip())

        # Overlay style
        self.font_scale = QtWidgets.QDoubleSpinBox()
        self.font_scale.setRange(0.5, 3.0)
        self.font_scale.setSingleStep(0.1)
        self.font_scale.setValue(float(getattr(cfg, 'font_scale', 1.0) or 1.0))

        self.overlay_font_family = QtWidgets.QFontComboBox()
        try:
            self.overlay_font_family.setCurrentFont(QtGui.QFont(str(getattr(cfg, 'overlay_font_family', 'Segoe UI') or 'Segoe UI')))
        except Exception:
            pass

        self.overlay_font_size = QtWidgets.QDoubleSpinBox()
        self.overlay_font_size.setRange(6.0, 96.0)
        self.overlay_font_size.setSingleStep(1.0)
        try:
            self.overlay_font_size.setValue(float(getattr(cfg, 'overlay_font_size', 14.0) or 14.0))
        except Exception:
            self.overlay_font_size.setValue(14.0)

        self.overlay_padding = QtWidgets.QSpinBox()
        self.overlay_padding.setRange(0, 60)
        try:
            self.overlay_padding.setValue(int(getattr(cfg, 'overlay_padding', 6) or 6))
        except Exception:
            self.overlay_padding.setValue(6)

        self.overlay_round_radius = QtWidgets.QSpinBox()
        self.overlay_round_radius.setRange(0, 40)
        try:
            self.overlay_round_radius.setValue(int(getattr(cfg, 'overlay_round_radius', 8) or 8))
        except Exception:
            self.overlay_round_radius.setValue(8)

        self.overlay_bg_color = QtWidgets.QLineEdit(_rgb_to_hex(getattr(cfg, 'overlay_bg_color', [0, 0, 0]), '#000000'))
        self.overlay_bg_opacity = QtWidgets.QSpinBox()
        self.overlay_bg_opacity.setRange(0, 255)
        try:
            self.overlay_bg_opacity.setValue(int(getattr(cfg, 'overlay_bg_opacity', 170) or 170))
        except Exception:
            self.overlay_bg_opacity.setValue(170)

        self.overlay_text_color = QtWidgets.QLineEdit(_rgb_to_hex(getattr(cfg, 'overlay_text_color', [255, 255, 255]), '#FFFFFF'))
        self.overlay_text_opacity = QtWidgets.QSpinBox()
        self.overlay_text_opacity.setRange(0, 255)
        try:
            self.overlay_text_opacity.setValue(int(getattr(cfg, 'overlay_text_opacity', 235) or 235))
        except Exception:
            self.overlay_text_opacity.setValue(235)

        self.overlay_wrap_chars = QtWidgets.QSpinBox()
        self.overlay_wrap_chars.setRange(0, 500)
        self.overlay_wrap_chars.setSingleStep(5)
        try:
            self.overlay_wrap_chars.setValue(int(getattr(cfg, 'overlay_max_chars_per_line', 0) or 0))
        except Exception:
            self.overlay_wrap_chars.setValue(0)

        self.overlay_max_box_height = QtWidgets.QSpinBox()
        self.overlay_max_box_height.setRange(0, 4000)
        self.overlay_max_box_height.setSingleStep(20)
        try:
            self.overlay_max_box_height.setValue(int(getattr(cfg, 'overlay_max_box_height', 0) or 0))
        except Exception:
            self.overlay_max_box_height.setValue(0)

        self.overlay_use_ocr_bg_color = QtWidgets.QCheckBox('Використовувати колір фону з OCR (якщо доступний)')
        self.overlay_use_ocr_bg_color.setChecked(bool(getattr(cfg, 'overlay_use_ocr_bg_color', False)))

        # Continuous
        self.continuous_mode = QtWidgets.QCheckBox('Увімкнено')
        self.continuous_mode.setChecked(bool(getattr(cfg, 'continuous_mode', False)))

        self.continuous_interval = QtWidgets.QSpinBox()
        self.continuous_interval.setRange(300, 10000)
        self.continuous_interval.setSingleStep(200)
        self.continuous_interval.setValue(int(getattr(cfg, 'continuous_interval_ms', 2000) or 2000))

        # Hotkeys
        self.hk_translate = HotkeyEditor(self)
        self.hk_translate.set_value(
            bool(getattr(cfg, 'hotkey_mod_ctrl', True)),
            bool(getattr(cfg, 'hotkey_mod_shift', True)),
            bool(getattr(cfg, 'hotkey_mod_alt', False)),
            int(getattr(cfg, 'hotkey_vk', 0x54) or 0x54),
        )

        self.hk_overlay = HotkeyEditor(self)
        self.hk_overlay.set_value(
            bool(getattr(cfg, 'overlay_hotkey_mod_ctrl', True)),
            bool(getattr(cfg, 'overlay_hotkey_mod_shift', True)),
            bool(getattr(cfg, 'overlay_hotkey_mod_alt', False)),
            int(getattr(cfg, 'overlay_hotkey_vk', 0x4F) or 0x4F),
        )

        self.hk_roi_add = HotkeyEditor(self)
        self.hk_roi_add.set_value(
            bool(getattr(cfg, 'roi_add_hotkey_mod_ctrl', True)),
            bool(getattr(cfg, 'roi_add_hotkey_mod_shift', True)),
            bool(getattr(cfg, 'roi_add_hotkey_mod_alt', False)),
            int(getattr(cfg, 'roi_add_hotkey_vk', 0x4E) or 0x4E),
        )

        self.hk_roi_manage = HotkeyEditor(self)
        self.hk_roi_manage.set_value(
            bool(getattr(cfg, 'roi_manage_hotkey_mod_ctrl', True)),
            bool(getattr(cfg, 'roi_manage_hotkey_mod_shift', True)),
            bool(getattr(cfg, 'roi_manage_hotkey_mod_alt', False)),
            int(getattr(cfg, 'roi_manage_hotkey_vk', 0x52) or 0x52),
        )

        self.hk_settings = HotkeyEditor(self)
        self.hk_settings.set_value(
            bool(getattr(cfg, 'settings_hotkey_mod_ctrl', True)),
            bool(getattr(cfg, 'settings_hotkey_mod_shift', True)),
            bool(getattr(cfg, 'settings_hotkey_mod_alt', False)),
            int(getattr(cfg, 'settings_hotkey_vk', 0x53) or 0x53),
        )

        self.hk_target_window = HotkeyEditor(self)
        self.hk_target_window.set_value(
            bool(getattr(cfg, 'target_window_hotkey_mod_ctrl', True)),
            bool(getattr(cfg, 'target_window_hotkey_mod_shift', True)),
            bool(getattr(cfg, 'target_window_hotkey_mod_alt', False)),
            int(getattr(cfg, 'target_window_hotkey_vk', 0x57) or 0x57),
        )

        self.hk_continuous_off = HotkeyEditor(self)
        self.hk_continuous_off.set_value(
            bool(getattr(cfg, 'continuous_off_hotkey_mod_ctrl', True)),
            bool(getattr(cfg, 'continuous_off_hotkey_mod_shift', True)),
            bool(getattr(cfg, 'continuous_off_hotkey_mod_alt', False)),
            int(getattr(cfg, 'continuous_off_hotkey_vk', 0x58) or 0x58),
        )

        # --- layout (grid of categories) ---
        root = QtWidgets.QVBoxLayout(self)

        scroll = QtWidgets.QScrollArea(self)
        scroll.setWidgetResizable(True)

        content = QtWidgets.QWidget(self)
        scroll.setWidget(content)

        grid = QtWidgets.QGridLayout(content)
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(14)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)

        def _group(title: str) -> tuple[QtWidgets.QGroupBox, QtWidgets.QFormLayout]:
            gb = QtWidgets.QGroupBox(title, content)
            fl = QtWidgets.QFormLayout(gb)
            fl.setLabelAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)
            fl.setFormAlignment(QtCore.Qt.AlignmentFlag.AlignTop)
            fl.setFieldGrowthPolicy(QtWidgets.QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
            return gb, fl

        gb_lang, f_lang = _group('Мови')
        f_lang.addRow('Мова джерела (auto/en/ja/…):', self.source_lang)
        f_lang.addRow('Мова перекладу (uk/en/…):', self.target_lang)

        gb_capture, f_cap = _group('Захоплення')
        f_cap.addRow('Монітор:', self.monitor)
        f_cap.addRow('Постійний режим:', self.continuous_mode)
        f_cap.addRow('Інтервал (мс):', self.continuous_interval)

        gb_ocr, f_ocr = _group('OCR та переклад')
        f_ocr.addRow('OCR провайдер:', self.ocr_provider)
        f_ocr.addRow('Провайдер перекладу:', self.translation_provider)
        f_ocr.addRow('Google Cloud API key (опційно):', self.google_api_key)
        f_ocr.addRow('OCR.space API key (опційно):', self.ocrspace_api_key)
        f_ocr.addRow(self.skip_translate_if_same_text)
        f_ocr.addRow(self.merge_close_text_regions)
        f_ocr.addRow(self.merge_vertical_lines)
        f_ocr.addRow(self.spellcheck_enabled)
        f_ocr.addRow('Мова для spellcheck (порожньо = source):', self.spellcheck_lang)

        gb_overlay, f_ov = _group('Оверлей')
        f_ov.addRow('Масштаб шрифту:', self.font_scale)
        f_ov.addRow('Шрифт:', self.overlay_font_family)
        f_ov.addRow('Розмір шрифту:', self.overlay_font_size)
        f_ov.addRow('Відступ у боксі (px):', self.overlay_padding)
        f_ov.addRow('Закруглення (px):', self.overlay_round_radius)
        f_ov.addRow('Колір боксу (hex):', self.overlay_bg_color)
        f_ov.addRow('Прозорість боксу (0-255):', self.overlay_bg_opacity)
        f_ov.addRow('Колір тексту (hex):', self.overlay_text_color)
        f_ov.addRow('Прозорість тексту (0-255):', self.overlay_text_opacity)
        f_ov.addRow('Переносити після N символів (0=word wrap):', self.overlay_wrap_chars)
        f_ov.addRow('Макс. висота бокса (0=без ліміту):', self.overlay_max_box_height)
        f_ov.addRow(self.overlay_use_ocr_bg_color)

        gb_hk, f_hk = _group('Гарячі клавіші')
        f_hk.addRow('Перекласти (toggle):', self.hk_translate)
        f_hk.addRow('Показ/ховати оверлей:', self.hk_overlay)
        f_hk.addRow('Додати ROI секцію:', self.hk_roi_add)
        f_hk.addRow('Відкрити ROI секції:', self.hk_roi_manage)
        f_hk.addRow('Відкрити налаштування:', self.hk_settings)
        f_hk.addRow('Вибрати target window під курсором:', self.hk_target_window)
        f_hk.addRow('Вимкнути постійний режим:', self.hk_continuous_off)

        grid.addWidget(gb_lang, 0, 0)
        grid.addWidget(gb_capture, 0, 1)
        grid.addWidget(gb_ocr, 1, 0)
        grid.addWidget(gb_overlay, 1, 1)
        grid.addWidget(gb_hk, 2, 0, 1, 2)

        root.addWidget(scroll, 1)

        btns = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Save | QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        root.addWidget(btns)

    def apply(self) -> None:
        self.cfg.source_lang = self.source_lang.text().strip() or 'auto'
        self.cfg.target_lang = self.target_lang.text().strip() or 'uk'

        mon = self.monitor.currentData()
        if mon is None:
            mon = 0
        self.cfg.capture_monitor = int(mon)

        self.cfg.ocr_provider = self.ocr_provider.currentText()
        self.cfg.translation_provider = self.translation_provider.currentText()
        self.cfg.google_api_key = self.google_api_key.text().strip()
        self.cfg.ocrspace_api_key = self.ocrspace_api_key.text().strip() or 'helloworld'

        self.cfg.skip_translate_if_same_text = bool(self.skip_translate_if_same_text.isChecked())
        self.cfg.merge_close_text_regions = bool(self.merge_close_text_regions.isChecked())
        self.cfg.merge_vertical_lines = bool(self.merge_vertical_lines.isChecked())

        self.cfg.spellcheck_enabled = bool(self.spellcheck_enabled.isChecked())
        self.cfg.spellcheck_lang = self.spellcheck_lang.text().strip()

        self.cfg.font_scale = float(self.font_scale.value())

        # Overlay style
        try:
            self.cfg.overlay_font_family = str(self.overlay_font_family.currentFont().family() or 'Segoe UI')
        except Exception:
            self.cfg.overlay_font_family = 'Segoe UI'

        try:
            self.cfg.overlay_font_size = float(self.overlay_font_size.value())
        except Exception:
            self.cfg.overlay_font_size = 14.0

        try:
            self.cfg.overlay_padding = int(self.overlay_padding.value())
        except Exception:
            self.cfg.overlay_padding = 6

        try:
            self.cfg.overlay_round_radius = int(self.overlay_round_radius.value())
        except Exception:
            self.cfg.overlay_round_radius = 8

        self.cfg.overlay_bg_color = _hex_to_rgb(self.overlay_bg_color.text(), [0, 0, 0])
        try:
            self.cfg.overlay_bg_opacity = int(self.overlay_bg_opacity.value())
        except Exception:
            self.cfg.overlay_bg_opacity = 170

        self.cfg.overlay_text_color = _hex_to_rgb(self.overlay_text_color.text(), [255, 255, 255])
        try:
            self.cfg.overlay_text_opacity = int(self.overlay_text_opacity.value())
        except Exception:
            self.cfg.overlay_text_opacity = 235

        try:
            self.cfg.overlay_max_chars_per_line = int(self.overlay_wrap_chars.value())
        except Exception:
            self.cfg.overlay_max_chars_per_line = 0

        try:
            self.cfg.overlay_max_box_height = int(self.overlay_max_box_height.value())
        except Exception:
            self.cfg.overlay_max_box_height = 0

        self.cfg.overlay_use_ocr_bg_color = bool(self.overlay_use_ocr_bg_color.isChecked())

        self.cfg.continuous_mode = bool(self.continuous_mode.isChecked())
        self.cfg.continuous_interval_ms = int(self.continuous_interval.value())

        # Hotkeys
        c, s, a, vk = self.hk_translate.get_value()
        self.cfg.hotkey_mod_ctrl = bool(c)
        self.cfg.hotkey_mod_shift = bool(s)
        self.cfg.hotkey_mod_alt = bool(a)
        self.cfg.hotkey_vk = int(vk)

        c, s, a, vk = self.hk_overlay.get_value()
        self.cfg.overlay_hotkey_mod_ctrl = bool(c)
        self.cfg.overlay_hotkey_mod_shift = bool(s)
        self.cfg.overlay_hotkey_mod_alt = bool(a)
        self.cfg.overlay_hotkey_vk = int(vk)

        c, s, a, vk = self.hk_roi_add.get_value()
        self.cfg.roi_add_hotkey_mod_ctrl = bool(c)
        self.cfg.roi_add_hotkey_mod_shift = bool(s)
        self.cfg.roi_add_hotkey_mod_alt = bool(a)
        self.cfg.roi_add_hotkey_vk = int(vk)

        c, s, a, vk = self.hk_roi_manage.get_value()
        self.cfg.roi_manage_hotkey_mod_ctrl = bool(c)
        self.cfg.roi_manage_hotkey_mod_shift = bool(s)
        self.cfg.roi_manage_hotkey_mod_alt = bool(a)
        self.cfg.roi_manage_hotkey_vk = int(vk)

        c, s, a, vk = self.hk_settings.get_value()
        self.cfg.settings_hotkey_mod_ctrl = bool(c)
        self.cfg.settings_hotkey_mod_shift = bool(s)
        self.cfg.settings_hotkey_mod_alt = bool(a)
        self.cfg.settings_hotkey_vk = int(vk)

        c, s, a, vk = self.hk_target_window.get_value()
        self.cfg.target_window_hotkey_mod_ctrl = bool(c)
        self.cfg.target_window_hotkey_mod_shift = bool(s)
        self.cfg.target_window_hotkey_mod_alt = bool(a)
        self.cfg.target_window_hotkey_vk = int(vk)

        c, s, a, vk = self.hk_continuous_off.get_value()
        self.cfg.continuous_off_hotkey_mod_ctrl = bool(c)
        self.cfg.continuous_off_hotkey_mod_shift = bool(s)
        self.cfg.continuous_off_hotkey_mod_alt = bool(a)
        self.cfg.continuous_off_hotkey_vk = int(vk)

class AppController(QtCore.QObject):
    def __init__(self, app: QtWidgets.QApplication):
        super().__init__()
        self.app = app
        self.app.setQuitOnLastWindowClosed(False)
        self.cfg = load_config()

        self.overlay = OverlayWindow()
        self.overlay.hide()

        # Target window (HWND) для захоплення лише одного додатку
        self._target_hwnd: Optional[int] = None
        self._target_title: str = ''

        # Якщо можливо, виключаємо overlay з capture, щоб не треба було ховати його перед OCR.
        try:
            self.overlay.set_exclude_from_capture(bool(getattr(self.cfg, 'exclude_overlay_from_capture', True)))
        except Exception:
            pass

        # Оверлей можна вимикати окремо (корисно у continuous режимі)
        self._overlay_enabled = True
        self._last_cap = None
        self._last_items = None
        self._last_ocr_signature: Optional[str] = None
        self._last_translations: Optional[List[str]] = None

        # ROI секції: кожна працює окремо (свій таймер/оверлей/стан)
        self._roi_sections: dict[str, RoiSectionController] = {}
        self._sections_active = False

        self.signals = _WorkerSignals()
        self.signals.finished.connect(self._on_translate_done)
        self.signals.failed.connect(self._on_translate_failed)

        self._busy = False

        # Continuous mode
        self._continuous_active = False
        self._cont_timer = QtCore.QTimer(self)
        self._cont_timer.setTimerType(QtCore.Qt.TimerType.PreciseTimer)
        self._cont_timer.timeout.connect(self._tick_continuous)

        # Tray
        self.tray = QtWidgets.QSystemTrayIcon(self._make_icon(), self.app)
        self.tray.setToolTip('Win Screen Translator')
        menu = QtWidgets.QMenu()
        act_translate = menu.addAction('Перекласти (toggle)')
        act_translate.triggered.connect(self.toggle_translation)

        self.act_overlay = menu.addAction('Оверлей увімкнено')
        self.act_overlay.setCheckable(True)
        self.act_overlay.setChecked(True)
        self.act_overlay.toggled.connect(self._on_overlay_toggled)

        self.act_continuous = menu.addAction('Постійний переклад')
        self.act_continuous.setCheckable(True)
        self.act_continuous.setChecked(bool(self.cfg.continuous_mode))
        self.act_continuous.toggled.connect(self._on_continuous_setting_toggled)

        act_sections = menu.addAction('Секції (ROI)…')
        act_sections.triggered.connect(self.open_roi_sections)

        act_settings = menu.addAction('Налаштування…')
        act_settings.triggered.connect(self.open_settings)

        menu.addSeparator()

        # Target window submenu
        self.target_menu = menu.addMenu('Target window')
        self.act_target_status = self.target_menu.addAction('Не вибрано')
        self.act_target_status.setEnabled(False)

        act_pick_target = self.target_menu.addAction('Вибрати під курсором')
        act_pick_target.triggered.connect(self.pick_target_window_under_cursor)

        self.act_clear_target = self.target_menu.addAction('Очистити')
        self.act_clear_target.triggered.connect(self.clear_target_window)
        self.act_clear_target.setEnabled(False)

        menu.addSeparator()
        act_quit = menu.addAction('Вихід')
        act_quit.triggered.connect(self.quit)
        self.tray.setContextMenu(menu)
        self.tray.show()

        # Hotkeys
        self._setup_hotkeys()

        # Піднімаємо ROI секції з конфіга
        self._rebuild_roi_sections()
        self._apply_overlay_enabled_to_roi()
        self._apply_overlay_style_to_all()

        # Спробувати виключити ROI оверлеї з capture теж
        try:
            self._apply_capture_exclusion_to_all_overlays()
        except Exception:
            pass

        # Синхронізувати UI target window
        try:
            self._update_target_window_ui()
        except Exception:
            pass


    def _stop_hotkeys(self) -> None:
        for hk in [
            getattr(self, 'hotkey', None),
            getattr(self, 'hotkey_overlay', None),
            getattr(self, 'hotkey_roi_add', None),
            getattr(self, 'hotkey_roi_manage', None),
            getattr(self, 'hotkey_settings', None),
            getattr(self, 'hotkey_target_window', None),
            getattr(self, 'hotkey_continuous_off', None),
        ]:
            if hk is None:
                continue
            try:
                hk.stop()
            except Exception:
                pass

        self.hotkey = None
        self.hotkey_overlay = None
        self.hotkey_roi_add = None
        self.hotkey_roi_manage = None
        self.hotkey_settings = None
        self.hotkey_target_window = None
        self.hotkey_continuous_off = None
        self._overlay_hotkey_is_main = False

    def _setup_hotkeys(self, show_warnings: bool = True) -> None:
        # спершу зупиняємо старі
        try:
            self._stop_hotkeys()
        except Exception:
            pass

        spec_main = None
        spec_overlay = None

        # 1) Основний хоткей (переклад)
        try:
            spec_main = HotkeySpec(
                ctrl=bool(getattr(self.cfg, 'hotkey_mod_ctrl', True)),
                shift=bool(getattr(self.cfg, 'hotkey_mod_shift', True)),
                alt=bool(getattr(self.cfg, 'hotkey_mod_alt', False)),
                vk=int(getattr(self.cfg, 'hotkey_vk', 0) or 0),
            )

            if int(spec_main.vk) != 0:
                self.hotkey = GlobalHotkey(1, spec_main)
                self.hotkey.triggered.connect(self.toggle_translation)
                self.hotkey.start()
        except Exception as e:
            self.hotkey = None
            if show_warnings:
                self.tray.showMessage('Win Screen Translator', f'Гаряча клавіша не активна: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)

        # 2) Хоткей оверлею
        try:
            spec_overlay = HotkeySpec(
                ctrl=bool(getattr(self.cfg, 'overlay_hotkey_mod_ctrl', True)),
                shift=bool(getattr(self.cfg, 'overlay_hotkey_mod_shift', True)),
                alt=bool(getattr(self.cfg, 'overlay_hotkey_mod_alt', False)),
                vk=int(getattr(self.cfg, 'overlay_hotkey_vk', 0) or 0),
            )

            same_as_main = False
            if spec_main is not None:
                if spec_overlay.ctrl == spec_main.ctrl and spec_overlay.shift == spec_main.shift and spec_overlay.alt == spec_main.alt and int(spec_overlay.vk) == int(spec_main.vk):
                    same_as_main = True

            if same_as_main:
                self._overlay_hotkey_is_main = True
                self.hotkey_overlay = None
            else:
                if int(spec_overlay.vk) != 0:
                    self.hotkey_overlay = GlobalHotkey(2, spec_overlay)
                    self.hotkey_overlay.triggered.connect(self.toggle_overlay_visibility)
                    self.hotkey_overlay.start()
        except Exception as e:
            self.hotkey_overlay = None
            if show_warnings:
                self.tray.showMessage('Win Screen Translator', f'Хоткей оверлею не активний: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)

        # Helper: перевірка колізій
        def _same(a: HotkeySpec, b: HotkeySpec) -> bool:
            if a is None or b is None:
                return False
            if bool(a.ctrl) != bool(b.ctrl):
                return False
            if bool(a.shift) != bool(b.shift):
                return False
            if bool(a.alt) != bool(b.alt):
                return False
            if int(a.vk) != int(b.vk):
                return False
            return True

        spec_roi_add = None
        # 3) Хоткей швидкого додавання ROI
        try:
            spec_roi_add = HotkeySpec(
                ctrl=bool(getattr(self.cfg, 'roi_add_hotkey_mod_ctrl', True)),
                shift=bool(getattr(self.cfg, 'roi_add_hotkey_mod_shift', True)),
                alt=bool(getattr(self.cfg, 'roi_add_hotkey_mod_alt', False)),
                vk=int(getattr(self.cfg, 'roi_add_hotkey_vk', 0) or 0),
            )

            collision = False
            if spec_main is not None and _same(spec_roi_add, spec_main):
                collision = True
            if not collision and spec_overlay is not None and _same(spec_roi_add, spec_overlay):
                collision = True

            if not collision and int(spec_roi_add.vk) != 0:
                self.hotkey_roi_add = GlobalHotkey(3, spec_roi_add)
                self.hotkey_roi_add.triggered.connect(self.quick_add_roi_section)
                self.hotkey_roi_add.start()
        except Exception as e:
            self.hotkey_roi_add = None
            if show_warnings:
                self.tray.showMessage('Win Screen Translator', f'Хоткей ROI додавання не активний: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)

        spec_roi_manage = None
        # 4) Хоткей відкриття ROI секцій
        try:
            spec_roi_manage = HotkeySpec(
                ctrl=bool(getattr(self.cfg, 'roi_manage_hotkey_mod_ctrl', True)),
                shift=bool(getattr(self.cfg, 'roi_manage_hotkey_mod_shift', True)),
                alt=bool(getattr(self.cfg, 'roi_manage_hotkey_mod_alt', False)),
                vk=int(getattr(self.cfg, 'roi_manage_hotkey_vk', 0) or 0),
            )

            collision = False
            if spec_main is not None and _same(spec_roi_manage, spec_main):
                collision = True
            if not collision and spec_overlay is not None and _same(spec_roi_manage, spec_overlay):
                collision = True
            if not collision and spec_roi_add is not None and _same(spec_roi_manage, spec_roi_add):
                collision = True

            if not collision and int(spec_roi_manage.vk) != 0:
                self.hotkey_roi_manage = GlobalHotkey(4, spec_roi_manage)
                self.hotkey_roi_manage.triggered.connect(self.open_roi_sections)
                self.hotkey_roi_manage.start()
        except Exception as e:
            self.hotkey_roi_manage = None
            if show_warnings:
                self.tray.showMessage('Win Screen Translator', f'Хоткей ROI секцій не активний: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)

        # 5) Хоткей відкриття налаштувань
        spec_settings = None
        try:
            spec_settings = HotkeySpec(
                ctrl=bool(getattr(self.cfg, 'settings_hotkey_mod_ctrl', True)),
                shift=bool(getattr(self.cfg, 'settings_hotkey_mod_shift', True)),
                alt=bool(getattr(self.cfg, 'settings_hotkey_mod_alt', False)),
                vk=int(getattr(self.cfg, 'settings_hotkey_vk', 0) or 0),
            )

            collision = False
            if spec_main is not None and _same(spec_settings, spec_main):
                collision = True
            if not collision and spec_overlay is not None and _same(spec_settings, spec_overlay):
                collision = True
            if not collision and spec_roi_add is not None and _same(spec_settings, spec_roi_add):
                collision = True
            if not collision and spec_roi_manage is not None and _same(spec_settings, spec_roi_manage):
                collision = True

            if not collision and int(spec_settings.vk) != 0:
                self.hotkey_settings = GlobalHotkey(5, spec_settings)
                self.hotkey_settings.triggered.connect(self.open_settings)
                self.hotkey_settings.start()
        except Exception as e:
            self.hotkey_settings = None
            if show_warnings:
                self.tray.showMessage('Win Screen Translator', f'Хоткей налаштувань не активний: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)

        # 6) Хоткей вибору target window під курсором
        try:
            spec_target = HotkeySpec(
                ctrl=bool(getattr(self.cfg, 'target_window_hotkey_mod_ctrl', True)),
                shift=bool(getattr(self.cfg, 'target_window_hotkey_mod_shift', True)),
                alt=bool(getattr(self.cfg, 'target_window_hotkey_mod_alt', False)),
                vk=int(getattr(self.cfg, 'target_window_hotkey_vk', 0) or 0),
            )

            collision = False
            if spec_main is not None and _same(spec_target, spec_main):
                collision = True
            if not collision and spec_overlay is not None and _same(spec_target, spec_overlay):
                collision = True
            if not collision and spec_roi_add is not None and _same(spec_target, spec_roi_add):
                collision = True
            if not collision and spec_roi_manage is not None and _same(spec_target, spec_roi_manage):
                collision = True
            if not collision and spec_settings is not None and _same(spec_target, spec_settings):
                collision = True

            if not collision and int(spec_target.vk) != 0:
                self.hotkey_target_window = GlobalHotkey(6, spec_target)
                self.hotkey_target_window.triggered.connect(self.pick_target_window_under_cursor)
                self.hotkey_target_window.start()
        except Exception as e:
            self.hotkey_target_window = None
            if show_warnings:
                self.tray.showMessage('Win Screen Translator', f'Хоткей target window не активний: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)

        # 7) Хоткей вимкнення постійного режиму
        try:
            spec_cont_off = HotkeySpec(
                ctrl=bool(getattr(self.cfg, 'continuous_off_hotkey_mod_ctrl', True)),
                shift=bool(getattr(self.cfg, 'continuous_off_hotkey_mod_shift', True)),
                alt=bool(getattr(self.cfg, 'continuous_off_hotkey_mod_alt', False)),
                vk=int(getattr(self.cfg, 'continuous_off_hotkey_vk', 0) or 0),
            )

            collision = False
            if spec_main is not None and _same(spec_cont_off, spec_main):
                collision = True
            if not collision and spec_overlay is not None and _same(spec_cont_off, spec_overlay):
                collision = True
            if not collision and spec_roi_add is not None and _same(spec_cont_off, spec_roi_add):
                collision = True
            if not collision and spec_roi_manage is not None and _same(spec_cont_off, spec_roi_manage):
                collision = True
            if not collision and spec_settings is not None and _same(spec_cont_off, spec_settings):
                collision = True
            try:
                if not collision and spec_target is not None and _same(spec_cont_off, spec_target):
                    collision = True
            except Exception:
                pass

            if not collision and int(spec_cont_off.vk) != 0:
                self.hotkey_continuous_off = GlobalHotkey(7, spec_cont_off)
                self.hotkey_continuous_off.triggered.connect(self.disable_continuous_mode)
                self.hotkey_continuous_off.start()
        except Exception as e:
            self.hotkey_continuous_off = None
            if show_warnings:
                self.tray.showMessage('Win Screen Translator', f'Хоткей вимкнення постійного режиму не активний: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)

    def _has_roi_sections(self) -> bool:
        try:
            return bool(self.cfg.roi_sections) and int(len(self.cfg.roi_sections)) > 0
        except Exception:
            return False

    def _apply_overlay_enabled_to_roi(self) -> None:
        for c in (self._roi_sections or {}).values():
            try:
                c.set_overlay_enabled_global(self._overlay_enabled)
            except Exception:
                continue

    def _build_overlay_style(self) -> OverlayStyle:
        style = OverlayStyle()

        style.font_family = str(getattr(self.cfg, 'overlay_font_family', 'Segoe UI') or 'Segoe UI')
        try:
            style.font_size = float(getattr(self.cfg, 'overlay_font_size', 14.0) or 14.0)
        except Exception:
            style.font_size = 14.0

        try:
            style.padding = int(getattr(self.cfg, 'overlay_padding', 6) or 6)
        except Exception:
            style.padding = 6

        try:
            style.round_radius = int(getattr(self.cfg, 'overlay_round_radius', 8) or 8)
        except Exception:
            style.round_radius = 8

        style.bg_color = _coerce_rgb(getattr(self.cfg, 'overlay_bg_color', [0, 0, 0]), [0, 0, 0])
        try:
            style.bg_opacity = int(getattr(self.cfg, 'overlay_bg_opacity', 170) or 170)
        except Exception:
            style.bg_opacity = 170

        style.text_color = _coerce_rgb(getattr(self.cfg, 'overlay_text_color', [255, 255, 255]), [255, 255, 255])
        try:
            style.text_opacity = int(getattr(self.cfg, 'overlay_text_opacity', 235) or 235)
        except Exception:
            style.text_opacity = 235

        style.use_ocr_bg_color = bool(getattr(self.cfg, 'overlay_use_ocr_bg_color', False))
        try:
            style.max_chars_per_line = int(getattr(self.cfg, 'overlay_max_chars_per_line', 0) or 0)
        except Exception:
            style.max_chars_per_line = 0

        try:
            style.max_box_height = int(getattr(self.cfg, 'overlay_max_box_height', 0) or 0)
        except Exception:
            style.max_box_height = 0

        return style

    def _apply_overlay_style_to_all(self) -> None:
        style = self._build_overlay_style()

        try:
            self.overlay.set_style(style)
        except Exception:
            pass

        for c in (self._roi_sections or {}).values():
            try:
                c.set_overlay_style(style)
            except Exception:
                continue

    def _apply_capture_exclusion_to_all_overlays(self) -> None:
        enabled = bool(getattr(self.cfg, 'exclude_overlay_from_capture', True))
        try:
            self.overlay.set_exclude_from_capture(enabled)
        except Exception:
            pass

        for c in (self._roi_sections or {}).values():
            try:
                c.overlay.set_exclude_from_capture(enabled)
            except Exception:
                continue

    def _should_hide_overlays_before_capture(self) -> bool:
        """Вирішує, чи треба ховати оверлеї перед capture.

        Якщо Windows підтримує exclude-from-capture і воно реально застосувалось, ховати не треба.
        """
        if not bool(getattr(self.cfg, 'exclude_overlay_from_capture', True)):
            return True
        try:
            if not bool(self.overlay.is_excluded_from_capture()):
                return True
        except Exception:
            return True
        return False

    def _update_target_window_ui(self) -> None:
        title = str(self._target_title or '').strip()
        hwnd_ok = False
        if self._target_hwnd is not None and is_window_valid(self._target_hwnd):
            hwnd_ok = True
        if not hwnd_ok:
            self._target_hwnd = None
            self._target_title = ''
            title = ''

        if hasattr(self, 'act_target_status') and self.act_target_status is not None:
            if title:
                shown = title
                if len(shown) > 60:
                    shown = shown[:60] + '…'
                self.act_target_status.setText(f'Поточне: {shown}')
            else:
                self.act_target_status.setText('Не вибрано')

        if hasattr(self, 'act_clear_target') and self.act_clear_target is not None:
            self.act_clear_target.setEnabled(bool(title))

    @QtCore.Slot()
    def pick_target_window_under_cursor(self) -> None:
        # У continuous режимі оверлей часто під курсором і заважає WindowFromPoint.
        # Тому тимчасово ховаємо всі оверлеї, вибираємо вікно і повертаємо показ назад.
        was_main_visible = False
        try:
            was_main_visible = bool(self.overlay.isVisible())
        except Exception:
            was_main_visible = False

        # Ховаємо, щоб не зловити власний overlay
        try:
            self._hide_all_overlays_now()
        except Exception:
            pass

        hwnd, title = get_window_under_cursor()
        if hwnd is None or not is_window_valid(hwnd):
            # Повернемо показ, якщо він був
            if was_main_visible and self._overlay_enabled and self._last_cap is not None and self._last_items:
                try:
                    self._show_overlay(self._last_cap, self._last_items)
                except Exception:
                    pass

            self.tray.showMessage('Win Screen Translator', 'Не вдалося визначити вікно під курсором.', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)
            return

        title = get_window_title(int(hwnd))
        self._target_hwnd = int(hwnd)
        self._target_title = str(title or '').strip()
        self._update_target_window_ui()

        # Скидаємо кеш, щоб точно не показати старий результат через skip_translate_if_same_text
        self._last_ocr_signature = None
        self._last_translations = None

        msg = 'Target window встановлено.'
        if self._target_title:
            msg = f'Target window: {self._target_title}'
        self.tray.showMessage('Win Screen Translator', msg, QtWidgets.QSystemTrayIcon.MessageIcon.Information)

        # Якщо continuous активний (і не ROI), зробимо оновлення одразу
        try:
            if self._continuous_active and not self._has_roi_sections():
                self._tick_continuous()
        except Exception:
            pass

        # Повернемо показ оверлею, якщо він був
        if was_main_visible and self._overlay_enabled and self._last_cap is not None and self._last_items:
            try:
                self._show_overlay(self._last_cap, self._last_items)
            except Exception:
                pass

    @QtCore.Slot()
    def clear_target_window(self) -> None:
        self._target_hwnd = None
        self._target_title = ''
        self._update_target_window_ui()
        self.tray.showMessage('Win Screen Translator', 'Target window очищено.', QtWidgets.QSystemTrayIcon.MessageIcon.Information)

    def _rebuild_roi_sections(self) -> None:
        # зупиняємо старі
        for c in (self._roi_sections or {}).values():
            try:
                c.stop()
            except Exception:
                pass

        self._roi_sections = {}

        roi_list = self.cfg.roi_sections or []
        out: List[RoiSectionConfig] = []
        for s in roi_list:
            if isinstance(s, dict):
                try:
                    sec = RoiSectionConfig.from_dict(s)
                except Exception:
                    continue
            else:
                sec = s
            out.append(sec)

        self.cfg.roi_sections = out

        for sec in out:
            try:
                ctrl = RoiSectionController(self.app, self.cfg, sec)
                self._roi_sections[sec.id] = ctrl
            except Exception as e:
                logger.error('Не вдалося створити ROI секцію %s: %s', getattr(sec, 'name', '?'), e)
                continue

    def _make_icon(self) -> QtGui.QIcon:
        # без файлів: стандартна іконка
        return self.app.style().standardIcon(QtWidgets.QStyle.StandardPixmap.SP_FileDialogInfoView)

    @QtCore.Slot()
    def toggle_translation(self) -> None:
        if self._busy:
            return

        # Якщо в налаштуваннях увімкнений постійний режим:
        #  - якщо continuous вже активний і оверлей-хоткей = основному -> хоткей ховає/показує оверлей
        #  - інакше хоткей працює як Start/Stop.
        if bool(getattr(self.cfg, 'continuous_mode', False)):
            # ROI секції: в continuous режимі головний хоткей = Start/Stop секцій.
            # Якщо хоткей оверлею збігається з головним і секції вже активні,
            # тоді головний хоткей лише ховає/показує оверлеї.
            if self._has_roi_sections():
                if self._sections_active and bool(getattr(self, '_overlay_hotkey_is_main', False)):
                    self.toggle_overlay_visibility()
                    return

                if self._sections_active:
                    self.stop_continuous()
                else:
                    self.start_continuous()
                return

            # Старий режим (весь екран)
            if self._continuous_active and bool(getattr(self, '_overlay_hotkey_is_main', False)):
                self.toggle_overlay_visibility()
                return

            if self._continuous_active:
                self.stop_continuous()
            else:
                self.start_continuous()
            return

        # Manual режим
        if self._has_roi_sections():
            any_visible = False
            for c in (self._roi_sections or {}).values():
                try:
                    if c.overlay.isVisible():
                        any_visible = True
                        break
                except Exception:
                    continue

            if any_visible:
                for c in (self._roi_sections or {}).values():
                    try:
                        c.hide()
                    except Exception:
                        continue
                return

            self.tray.showMessage(
                'Win Screen Translator',
                'Оновлюю секції: OCR + переклад…',
                QtWidgets.QSystemTrayIcon.MessageIcon.Information,
            )

            # Запускаємо оновлення кожної секції незалежно
            for c in (self._roi_sections or {}).values():
                try:
                    c.update_once()
                except Exception:
                    continue
            return

        if self.overlay.isVisible():
            self.overlay.hide()
            return

        self._busy = True
        self.tray.showMessage('Win Screen Translator', 'Знімаю екран, роблю OCR і переклад…', QtWidgets.QSystemTrayIcon.MessageIcon.Information)

        if self._should_hide_overlays_before_capture():
            self._hide_all_overlays_now()
        worker = TranslateWorker(
            self.cfg,
            self.signals,
            target_hwnd=self._target_hwnd,
            prev_signature=self._last_ocr_signature,
            prev_translations=self._last_translations,
        )
        worker.start()

    @QtCore.Slot()
    def toggle_overlay_visibility(self) -> None:
        """Вмикає/вимикає показ оверлею, не чіпаючи continuous OCR."""
        self._set_overlay_enabled(not self._overlay_enabled)


    def _hide_all_overlays_now(self) -> None:
        """Ховає всі оверлеї перед захопленням екрану, щоб OCR не читав текст із попередніх боксів."""
        try:
            # Головний оверлей
            try:
                if self.overlay is not None:
                    self.overlay.hide()
            except Exception:
                pass

            # ROI оверлеї (якщо активні)
            for c in (self._roi_sections or {}).values():
                try:
                    c.hide()
                except Exception:
                    continue

            # Дати Qt/Windows один «тик», щоб compositing реально прибрав вікна з кадру
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

    def _set_overlay_enabled(self, enabled: bool) -> None:
        self._overlay_enabled = bool(enabled)

        # Синхронізуємо пункт меню
        try:
            if hasattr(self, 'act_overlay') and self.act_overlay is not None:
                self.act_overlay.blockSignals(True)
                self.act_overlay.setChecked(self._overlay_enabled)
                if self._overlay_enabled:
                    self.act_overlay.setText('Оверлей увімкнено')
                else:
                    self.act_overlay.setText('Оверлей вимкнено')
        finally:
            try:
                self.act_overlay.blockSignals(False)
            except Exception:
                pass

        self._apply_overlay_enabled_to_roi()

        if not self._overlay_enabled:
            self.overlay.hide()
            self.tray.showMessage('Win Screen Translator', 'Оверлей: OFF', QtWidgets.QSystemTrayIcon.MessageIcon.Information)
            return

        # Якщо є останній результат, показуємо його одразу
        if self._last_cap is not None and self._last_items:
            try:
                self._show_overlay(self._last_cap, self._last_items)
            except Exception:
                pass
        self.tray.showMessage('Win Screen Translator', 'Оверлей: ON', QtWidgets.QSystemTrayIcon.MessageIcon.Information)

    def _on_overlay_toggled(self, checked: bool) -> None:
        self._set_overlay_enabled(bool(checked))

    def _on_continuous_setting_toggled(self, checked: bool) -> None:
        # Це лише налаштування. Реальний старт/стоп робить toggle_translation або кнопка нижче.
        self.cfg.continuous_mode = bool(checked)
        try:
            save_config(self.cfg)
        except Exception as e:
            self.tray.showMessage('Win Screen Translator', f'Не вдалося зберегти налаштування: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)
            return

        # Якщо вимкнули, а воно зараз крутиться, зупиняємо.
        if not checked:
            if self._continuous_active:
                self.stop_continuous()
            if self._sections_active:
                self.stop_continuous()


    @QtCore.Slot()
    def disable_continuous_mode(self) -> None:
        """Повністю вимикає continuous_mode і зупиняє поточні таймери.

        Це окремий хоткей, бо в continuous режимі основний хоткей може бути зайнятий start/stop.
        """
        # Вимикаємо саме налаштування
        self.cfg.continuous_mode = False

        # Зупиняємо все, що може крутитись
        try:
            if self._continuous_active or self._sections_active:
                self.stop_continuous()
        except Exception:
            pass

        # Синхронізуємо меню
        try:
            if hasattr(self, 'act_continuous') and self.act_continuous is not None:
                self.act_continuous.blockSignals(True)
                self.act_continuous.setChecked(False)
        finally:
            try:
                if hasattr(self, 'act_continuous') and self.act_continuous is not None:
                    self.act_continuous.blockSignals(False)
            except Exception:
                pass

        try:
            save_config(self.cfg)
        except Exception as e:
            self.tray.showMessage('Win Screen Translator', f'Не вдалося зберегти налаштування: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)
            return

        self.tray.showMessage('Win Screen Translator', 'Постійний режим вимкнено', QtWidgets.QSystemTrayIcon.MessageIcon.Information)
    def start_continuous(self) -> None:
        if self._has_roi_sections():
            if self._sections_active:
                return
            self._sections_active = True

            started = 0
            for c in (self._roi_sections or {}).values():
                try:
                    c.start()
                    started += 1
                except Exception:
                    continue

            self.tray.showMessage(
                'Win Screen Translator',
                f'Постійний режим (ROI): ON (секцій: {started})',
                QtWidgets.QSystemTrayIcon.MessageIcon.Information,
            )
            return

        if self._continuous_active:
            return
        self._continuous_active = True

        interval = int(getattr(self.cfg, 'continuous_interval_ms', 2000))
        # Не даємо ставити зовсім божевільний інтервал для мережевих OCR.
        if self.cfg.ocr_provider == 'ocrspace' and interval < 1500:
            interval = 1500
        if interval < 300:
            interval = 300

        self._cont_timer.start(interval)
        self.tray.showMessage('Win Screen Translator', f'Постійний режим: ON ({interval} мс)', QtWidgets.QSystemTrayIcon.MessageIcon.Information)

        # Одразу робимо перший прохід
        self._tick_continuous()

    def stop_continuous(self) -> None:
        if self._has_roi_sections():
            if not self._sections_active:
                return
            self._sections_active = False

            for c in (self._roi_sections or {}).values():
                try:
                    c.stop()
                except Exception:
                    continue

            self.tray.showMessage(
                'Win Screen Translator',
                'Постійний режим (ROI): OFF',
                QtWidgets.QSystemTrayIcon.MessageIcon.Information,
            )
            return

        if not self._continuous_active:
            return
        self._continuous_active = False
        self._cont_timer.stop()
        self.overlay.hide()
        self.tray.showMessage('Win Screen Translator', 'Постійний режим: OFF', QtWidgets.QSystemTrayIcon.MessageIcon.Information)

    def _tick_continuous(self) -> None:
        if self._has_roi_sections():
            # ROI секції мають власні таймери.
            return
        if self._busy:
            return
        self._busy = True
        if self._should_hide_overlays_before_capture():
            self._hide_all_overlays_now()
        worker = TranslateWorker(
            self.cfg,
            self.signals,
            target_hwnd=self._target_hwnd,
            prev_signature=self._last_ocr_signature,
            prev_translations=self._last_translations,
        )
        worker.start()

    def _find_mss_monitor_for_point(self, x: int, y: int):
        """Повертає MSS-монітор (dict) для заданої фізичної точки (x,y).

        Якщо точка не потрапляє ні в один монітор (рідкісні DPI/virtual кейси),
        повертаємо найближчий за відстанню.
        """
        try:
            mons = list_monitors()
        except Exception:
            mons = []

        best = None
        best_dist = 10 ** 18

        for m in (mons or []):
            try:
                ml = int(m.get('left', 0))
                mt = int(m.get('top', 0))
                mw = int(m.get('width', 0))
                mh = int(m.get('height', 0))
                if mw <= 0 or mh <= 0:
                    continue

                mr = ml + mw
                mb = mt + mh

                inside = (int(x) >= ml) and (int(x) < mr) and (int(y) >= mt) and (int(y) < mb)
                if inside:
                    return m

                dx = 0
                if int(x) < ml:
                    dx = ml - int(x)
                elif int(x) > mr:
                    dx = int(x) - mr

                dy = 0
                if int(y) < mt:
                    dy = mt - int(y)
                elif int(y) > mb:
                    dy = int(y) - mb

                dist = int(dx) * int(dx) + int(dy) * int(dy)
                if dist < best_dist:
                    best_dist = dist
                    best = m
            except Exception:
                continue

        return best

    def _best_screen_for_capture(self, cap):
        """Підбирає QScreen для MSS-координат захоплення."""
        screens = self.app.screens() or []
        if not screens:
            return self.app.primaryScreen(), 1.0

        # Якщо захоплюємо конкретний монітор, намагаємось матчити по фізичних координатах.
        if getattr(cap, 'monitor_index', 0) != 0:
            ref_left = int(getattr(cap, 'left', 0) or 0)
            ref_top = int(getattr(cap, 'top', 0) or 0)
            ref_w = int(getattr(cap, 'width', 0) or 0)
            ref_h = int(getattr(cap, 'height', 0) or 0)

            # monitor_index == -1 означає захоплення target window.
            # Для підбору екрана матчимо не за розміром вікна, а за монітором, на якому воно знаходиться.
            if int(getattr(cap, 'monitor_index', 0) or 0) == -1:
                try:
                    mon = self._find_mss_monitor_for_point(ref_left + int(ref_w / 2), ref_top + int(ref_h / 2))
                    if mon is not None:
                        ref_left = int(mon.get('left', ref_left) or ref_left)
                        ref_top = int(mon.get('top', ref_top) or ref_top)
                        ref_w = int(mon.get('width', ref_w) or ref_w)
                        ref_h = int(mon.get('height', ref_h) or ref_h)
                except Exception:
                    pass

            best = None
            best_score = 10**18
            for s in screens:
                try:
                    g = s.geometry()  # DIPs
                    dpr = float(s.devicePixelRatio() or 1.0)
                    phys_left = int(round(g.left() * dpr))
                    phys_top = int(round(g.top() * dpr))
                    phys_w = int(round(g.width() * dpr))
                    phys_h = int(round(g.height() * dpr))
                    score = abs(phys_left - int(ref_left)) + abs(phys_top - int(ref_top)) + abs(phys_w - int(ref_w)) + abs(phys_h - int(ref_h))
                    if score < best_score:
                        best = s
                        best_score = score
                except Exception:
                    continue

            if best is None:
                best = self.app.primaryScreen()
            return best, float(best.devicePixelRatio() or 1.0)

        # Virtual screen: беремо primaryScreen і її virtualGeometry
        s = self.app.primaryScreen()
        return s, float(s.devicePixelRatio() or 1.0)

    @QtCore.Slot(object, object, object)
    def _on_translate_done(self, cap, items, meta):
        self._busy = False

        # Кешуємо, щоб можна було знову показати оверлей після вимкнення.
        self._last_cap = cap
        self._last_items = items

        try:
            if isinstance(meta, dict):
                self._last_ocr_signature = meta.get('signature')
                self._last_translations = meta.get('translations')
        except Exception:
            pass

        if not items:
            if self._overlay_enabled:
                self.overlay.hide()
            return

        # Якщо користувач встиг вимкнути постійний режим, поки йшов OCR,
        # просто ігноруємо результат.
        if bool(getattr(self.cfg, 'continuous_mode', False)) and not self._continuous_active:
            return

        if not self._overlay_enabled:
            return
        try:
            self._show_overlay(cap, items)
        except Exception as e:
            self.tray.showMessage('Win Screen Translator', f'Не вдалося показати оверлей: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Critical)

    def _show_overlay(self, cap: CaptureInfo, items: List[OverlayItem]) -> None:
        screen, dpr = self._best_screen_for_capture(cap)

        try:
            screen_name = ''
            if screen is not None:
                try:
                    screen_name = screen.name()
                except Exception:
                    screen_name = ''
            logger.info(
                "Overlay: monitor=%s screen=%s dpr=%.3f items=%d",
                getattr(cap, 'monitor_index', '?'),
                screen_name,
                float(dpr or 1.0),
                len(items),
            )
        except Exception:
            pass

        # Геометрія оверлею повинна бути в координатах Qt (DIP)
        if getattr(cap, 'monitor_index', 0) == 0:
            vg = screen.virtualGeometry()
            self.overlay.set_overlay_geometry(int(vg.left()), int(vg.top()), int(vg.width()), int(vg.height()))

            # Для virtual screen робимо просту пропорцію. Якщо у моніторів різний scaling,
            # краще вибрати конкретний монітор у налаштуваннях.
            sx = float(vg.width()) / float(max(1, int(cap.width)))
            sy = float(vg.height()) / float(max(1, int(cap.height)))

            items_fixed = []
            for it in items:
                items_fixed.append(
                    OverlayItem(
                        left=int(round(it.left * sx)),
                        top=int(round(it.top * sy)),
                        right=int(round(it.right * sx)),
                        bottom=int(round(it.bottom * sy)),
                        text=it.text,
                        bg_color=it.bg_color,
                    )
                )
            self.overlay.set_items(items_fixed, font_scale=self.cfg.font_scale)
        elif int(getattr(cap, 'monitor_index', 0) or 0) == -1:
            # Target window: items зараз відносно (0,0) вікна.
            # Треба перевести їх у координати монітора і тільки потім у Qt DIP.
            mon_left = 0
            mon_top = 0
            try:
                mon = self._find_mss_monitor_for_point(int(cap.left) + int(int(cap.width) / 2), int(cap.top) + int(int(cap.height) / 2))
                if mon is not None:
                    mon_left = int(mon.get('left', 0) or 0)
                    mon_top = int(mon.get('top', 0) or 0)
            except Exception:
                pass

            g = screen.geometry()  # DIP
            self.overlay.set_overlay_geometry(int(g.left()), int(g.top()), int(g.width()), int(g.height()))

            inv = 1.0 / float(dpr or 1.0)
            items_fixed = []
            for it in items:
                abs_l = int(cap.left) + int(it.left)
                abs_t = int(cap.top) + int(it.top)
                abs_r = int(cap.left) + int(it.right)
                abs_b = int(cap.top) + int(it.bottom)

                rel_l = abs_l - int(mon_left)
                rel_t = abs_t - int(mon_top)
                rel_r = abs_r - int(mon_left)
                rel_b = abs_b - int(mon_top)

                items_fixed.append(
                    OverlayItem(
                        left=int(round(rel_l * inv)),
                        top=int(round(rel_t * inv)),
                        right=int(round(rel_r * inv)),
                        bottom=int(round(rel_b * inv)),
                        text=it.text,
                        bg_color=it.bg_color,
                    )
                )

            self.overlay.set_items(items_fixed, font_scale=self.cfg.font_scale)
        else:
            g = screen.geometry()  # DIP
            self.overlay.set_overlay_geometry(int(g.left()), int(g.top()), int(g.width()), int(g.height()))

            # MSS координати = фізичні пікселі. Qt координати = DIP.
            items_fixed = []
            inv = 1.0 / float(dpr or 1.0)
            for it in items:
                items_fixed.append(
                    OverlayItem(
                        left=int(round(it.left * inv)),
                        top=int(round(it.top * inv)),
                        right=int(round(it.right * inv)),
                        bottom=int(round(it.bottom * inv)),
                        text=it.text,
                        bg_color=it.bg_color,
                    )
                )
            self.overlay.set_items(items_fixed, font_scale=self.cfg.font_scale)

        self.overlay.show()
        try:
            self.overlay.raise_()
        except Exception:
            pass

    @QtCore.Slot(str)
    def _on_translate_failed(self, msg: str):
        self._busy = False

        if bool(getattr(self.cfg, 'continuous_mode', False)) and not self._continuous_active:
            return
        self.tray.showMessage('Win Screen Translator', f'Помилка: {msg}', QtWidgets.QSystemTrayIcon.MessageIcon.Critical)

    @QtCore.Slot()
    def open_settings(self) -> None:
        dlg = SettingsDialog(self.cfg)
        if dlg.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            try:
                dlg.apply()
                save_config(self.cfg)
            except Exception as e:
                # Не валимо весь процес через запис у файл.
                self.tray.showMessage('Win Screen Translator', f'Не вдалося зберегти налаштування: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)
                return

            try:
                self._apply_overlay_style_to_all()
            except Exception:
                pass

            try:
                self._apply_capture_exclusion_to_all_overlays()
            except Exception:
                pass

            # Синхронізуємо чекбокс у меню
            try:
                self.act_continuous.blockSignals(True)
                self.act_continuous.setChecked(bool(self.cfg.continuous_mode))
            finally:
                self.act_continuous.blockSignals(False)

            try:
                self._setup_hotkeys(show_warnings=True)
            except Exception:
                pass

            self.tray.showMessage('Win Screen Translator', 'Збережено.', QtWidgets.QSystemTrayIcon.MessageIcon.Information)

    @QtCore.Slot()
    def quick_add_roi_section(self) -> None:
        # Визначаємо монітор для вибору. Якщо у захопленні стоїть virtual screen (0), беремо 1.
        mon = int(getattr(self.cfg, 'capture_monitor', 0) or 0)
        if mon <= 0:
            mon = 1

        roi = None
        try:
            roi = select_roi_on_monitor(self.app, mon)
        except Exception as e:
            self.tray.showMessage('Win Screen Translator', f'Не вдалося вибрати зону: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)
            return

        if not roi:
            return

        # Якщо секції вже крутяться, спочатку зупиняємо, щоб не залишити старі оверлеї/таймери.
        was_active = bool(self._sections_active)
        if was_active:
            try:
                self.stop_continuous()
            except Exception:
                pass

        try:
            sec = RoiSectionConfig()
            sec.ensure_id()
            sec.monitor_index = int(mon)
            sec.x = int(roi[0])
            sec.y = int(roi[1])
            sec.width = int(roi[2])
            sec.height = int(roi[3])
            sec.enabled = True

            # Пер-секційні налаштування оверлею: копіюємо глобальні,
            # але font size і max box height залишаємо в авто-режимі.
            try:
                sec.overlay_custom = True
                sec.overlay_font_family = str(getattr(self.cfg, 'overlay_font_family', 'Segoe UI') or 'Segoe UI')
                sec.overlay_padding = int(getattr(self.cfg, 'overlay_padding', 6) or 6)
                sec.overlay_round_radius = int(getattr(self.cfg, 'overlay_round_radius', 8) or 8)

                sec.overlay_bg_color = list(getattr(self.cfg, 'overlay_bg_color', [0, 0, 0]) or [0, 0, 0])
                sec.overlay_bg_opacity = int(getattr(self.cfg, 'overlay_bg_opacity', 170) or 170)

                sec.overlay_text_color = list(getattr(self.cfg, 'overlay_text_color', [255, 255, 255]) or [255, 255, 255])
                sec.overlay_text_opacity = int(getattr(self.cfg, 'overlay_text_opacity', 235) or 235)

                sec.overlay_use_ocr_bg_color = bool(getattr(self.cfg, 'overlay_use_ocr_bg_color', False))
                sec.overlay_max_chars_per_line = int(getattr(self.cfg, 'overlay_max_chars_per_line', 0) or 0)

                sec.overlay_font_size = None
                sec.overlay_max_box_height = None
            except Exception:
                pass

            # Назва + інтервал за замовчуванням
            try:
                n = int(len(self.cfg.roi_sections or [])) + 1
            except Exception:
                n = 1
            sec.name = f'Секція {n}'

            try:
                sec.interval_ms = int(getattr(self.cfg, 'continuous_interval_ms', 2000) or 2000)
            except Exception:
                sec.interval_ms = 2000

            if self.cfg.roi_sections is None:
                self.cfg.roi_sections = []
            self.cfg.roi_sections.append(sec)

            save_config(self.cfg)
            self._rebuild_roi_sections()
            self._apply_overlay_enabled_to_roi()
            self._apply_overlay_style_to_all()
        except Exception as e:
            self.tray.showMessage('Win Screen Translator', f'Не вдалося додати секцію: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)
            return

        if was_active:
            try:
                self.start_continuous()
            except Exception:
                pass

        self.tray.showMessage('Win Screen Translator', 'ROI секцію додано.', QtWidgets.QSystemTrayIcon.MessageIcon.Information)

    @QtCore.Slot()
    def open_roi_sections(self) -> None:
        dlg = RoiSectionsDialog(self.app, self.cfg)
        res = dlg.exec()
        if res != QtWidgets.QDialog.DialogCode.Accepted:
            return

        # Зберігаємо + перезбираємо секції
        was_active = bool(self._sections_active)
        if was_active:
            try:
                self.stop_continuous()
            except Exception:
                pass

        try:
            save_config(self.cfg)
        except Exception as e:
            self.tray.showMessage('Win Screen Translator', f'Не вдалося зберегти секції: {e}', QtWidgets.QSystemTrayIcon.MessageIcon.Warning)
            return

        try:
            self._rebuild_roi_sections()
            self._apply_overlay_enabled_to_roi()
            self._apply_overlay_style_to_all()
            self._apply_capture_exclusion_to_all_overlays()
        except Exception:
            pass

        if was_active:
            try:
                self.start_continuous()
            except Exception:
                pass

        self.tray.showMessage('Win Screen Translator', 'Секції збережено.', QtWidgets.QSystemTrayIcon.MessageIcon.Information)

    @QtCore.Slot()
    def quit(self) -> None:
        try:
            if self.hotkey:
                self.hotkey.stop()
        except Exception:
            pass

        try:
            if getattr(self, 'hotkey_overlay', None):
                self.hotkey_overlay.stop()
        except Exception:
            pass

        try:
            if getattr(self, 'hotkey_roi_add', None):
                self.hotkey_roi_add.stop()
        except Exception:
            pass

        try:
            if self._sections_active:
                self.stop_continuous()
        except Exception:
            pass

        try:
            if self._continuous_active:
                self.stop_continuous()
        except Exception:
            pass

        # на всякий
        for c in (self._roi_sections or {}).values():
            try:
                c.stop()
            except Exception:
                pass
        self.tray.hide()
        self.app.quit()
