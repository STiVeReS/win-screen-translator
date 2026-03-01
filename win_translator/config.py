import json
import os
import uuid
import threading
from dataclasses import dataclass, asdict, field
from typing import Any, Dict, List, Optional


_SAVE_LOCK = threading.Lock()


def _default_config_dir() -> str:
    appdata = os.environ.get('APPDATA')
    if appdata:
        return os.path.join(appdata, 'WinScreenTranslator')
    # fallback
    return os.path.join(os.path.expanduser('~'), '.winscreentranslator')


@dataclass
class AppConfig:
    source_lang: str = 'auto'
    target_lang: str = 'uk'

    # ocrspace | googlecloud | rapidocr
    ocr_provider: str = 'ocrspace'
    # freegoogle | googlecloud | argos
    translation_provider: str = 'freegoogle'

    google_api_key: str = ''
    ocrspace_api_key: str = 'helloworld'  # free demo key

    font_scale: float = 1.0

    # Стиль боксів оверлею (загальні налаштування для всіх оверлеїв)
    overlay_font_family: str = 'Segoe UI'
    overlay_font_size: float = 14.0
    overlay_padding: int = 6
    overlay_round_radius: int = 8

    # Колір боксів (RGB) + прозорість 0..255
    overlay_bg_color: List[int] = field(default_factory=lambda: [0, 0, 0])
    overlay_bg_opacity: int = 170

    # Колір тексту (RGB) + прозорість 0..255
    overlay_text_color: List[int] = field(default_factory=lambda: [255, 255, 255])
    overlay_text_opacity: int = 235

    # Якщо True, і OCR віддає bg_color, будемо використовувати його замість фіксованого overlay_bg_color
    overlay_use_ocr_bg_color: bool = False

    # 0 = не чіпати (Qt WordWrap), >0 = переносити примусово по N символів
    overlay_max_chars_per_line: int = 0
    # Максимальна висота одного бокса (в DIP/px). 0 = без ліміту
    overlay_max_box_height: int = 0

    # Постпроцес OCR: зліплювати близькі бокси в один рядок (щоб текст не ламався на шматки)
    merge_close_text_regions: bool = True
    # Максимальний горизонтальний розрив між боксами, у частках від медіанної висоти рядка.
    # Більше значення = агресивніше зліплювання.
    merge_x_gap_ratio: float = 1.25
    # Допуск по Y (для віднесення бокса до того ж рядка), у частках від медіанної висоти.
    merge_line_y_ratio: float = 0.70
    # Додатково зліплювати сусідні рядки в один мультилайн-бокс.
    merge_vertical_lines: bool = True


    # Spellcheck / виправлення OCR тексту перед перекладом (експериментально)
    spellcheck_enabled: bool = False
    # Порожньо = брати source_lang. Для pyspellchecker реально працює насамперед en.
    spellcheck_lang: str = ''

    # Якщо текст не змінився, пропускаємо повторний переклад (економить API/час)
    skip_translate_if_same_text: bool = True

    # Захоплення екрана
    # 0 = усі монітори (virtual screen), 1..N = конкретний монітор у порядку MSS
    capture_monitor: int = 0

    # Постійний режим перекладу (оновлення по таймеру)
    continuous_mode: bool = False
    continuous_interval_ms: int = 2000

    # Хоткей повного вимкнення постійного режиму (вимикає continuous_mode і зупиняє таймери)
    # Дефолт: Ctrl+Shift+X
    continuous_off_hotkey_mod_ctrl: bool = True
    continuous_off_hotkey_mod_shift: bool = True
    continuous_off_hotkey_mod_alt: bool = False
    continuous_off_hotkey_vk: int = 0x58  # 'X'

    # Hotkey: Ctrl+Shift+T
    hotkey_mod_ctrl: bool = True
    hotkey_mod_shift: bool = True
    hotkey_mod_alt: bool = False
    hotkey_vk: int = 0x54  # 'T'

    # RapidOCR (optional)
    rapidocr_models_dir: str = ''
    rapidocr_min_confidence: float = 0.5
    rapidocr_box_thresh: float = 0.5
    rapidocr_unclip_ratio: float = 1.6

    # Debug: писати читабельний дамп OCR у лог
    debug_ocr_log: bool = True

    # Якщо True, на Windows намагаємось виключити overlay з capture (щоб не треба було ховати бокси).
    # Якщо ОС/метод захоплення не підтримує, програма автоматично повернеться до «ховати перед capture».
    exclude_overlay_from_capture: bool = True

    # Окремий хоткей для вмикання/вимикання оверлею (не зупиняє OCR у continuous режимі)
    # Дефолт: Ctrl+Shift+O
    overlay_hotkey_mod_ctrl: bool = True
    overlay_hotkey_mod_shift: bool = True
    overlay_hotkey_mod_alt: bool = False
    overlay_hotkey_vk: int = 0x4F  # 'O'

    # Хоткей швидкого додавання ROI секції (Ctrl+Shift+N)
    roi_add_hotkey_mod_ctrl: bool = True
    roi_add_hotkey_mod_shift: bool = True
    roi_add_hotkey_mod_alt: bool = False
    roi_add_hotkey_vk: int = 0x4E  # 'N'

    # Хоткей відкриття вікна ROI секцій (Ctrl+Shift+R)
    roi_manage_hotkey_mod_ctrl: bool = True
    roi_manage_hotkey_mod_shift: bool = True
    roi_manage_hotkey_mod_alt: bool = False
    roi_manage_hotkey_vk: int = 0x52  # 'R'

    # Хоткей відкриття налаштувань (Ctrl+Shift+S)
    settings_hotkey_mod_ctrl: bool = True
    settings_hotkey_mod_shift: bool = True
    settings_hotkey_mod_alt: bool = False
    settings_hotkey_vk: int = 0x53  # 'S'

    # Хоткей вибору target window під курсором (Ctrl+Shift+W)
    target_window_hotkey_mod_ctrl: bool = True
    target_window_hotkey_mod_shift: bool = True
    target_window_hotkey_mod_alt: bool = False
    target_window_hotkey_vk: int = 0x57  # 'W'

    # ROI секції (окремі області екрану)
    roi_sections: List["RoiSectionConfig"] = field(default_factory=list)

    def config_dir(self) -> str:
        return _default_config_dir()

    def config_path(self) -> str:
        return os.path.join(self.config_dir(), 'settings.json')


@dataclass
class RoiSectionConfig:
    """Опис однієї ROI-секції.

    Координати зберігаємо відносно монітора (у фізичних пікселях MSS).
    Так секції не «відлітають», якщо користувач переставляє монітори у Windows.
    """

    id: str = ''
    name: str = 'Секція'
    target_app: str = ''
    monitor_index: int = 1
    x: int = 0
    y: int = 0
    width: int = 320
    height: int = 180
    enabled: bool = True
    interval_ms: int = 2000

    # Пер-секційні налаштування оверлею.
    # Якщо overlay_custom=False, секція використовує глобальні налаштування з AppConfig.
    overlay_custom: bool = True

    overlay_font_family: Optional[str] = None
    overlay_font_size: Optional[float] = None
    overlay_padding: Optional[int] = None
    overlay_round_radius: Optional[int] = None

    overlay_bg_color: Optional[List[int]] = None
    overlay_bg_opacity: Optional[int] = None

    overlay_text_color: Optional[List[int]] = None
    overlay_text_opacity: Optional[int] = None

    overlay_use_ocr_bg_color: Optional[bool] = None
    overlay_max_chars_per_line: Optional[int] = None
    # None = авто (визначається по OCR), 0 = без ліміту
    overlay_max_box_height: Optional[int] = None

    def ensure_id(self) -> None:
        if self.id and str(self.id).strip():
            return
        self.id = uuid.uuid4().hex[:10]

    @staticmethod
    def from_dict(data: Dict[str, Any]) -> "RoiSectionConfig":
        cfg = RoiSectionConfig()

        raw_keys = set()

        if isinstance(data, dict):
            raw_keys = set(data.keys())
            for k, v in data.items():
                if hasattr(cfg, k):
                    setattr(cfg, k, v)

        cfg.ensure_id()

        try:
            cfg.monitor_index = int(cfg.monitor_index or 1)
        except Exception:
            cfg.monitor_index = 1

        try:
            cfg.x = int(cfg.x or 0)
            cfg.y = int(cfg.y or 0)
            cfg.width = int(cfg.width or 1)
            cfg.height = int(cfg.height or 1)
            cfg.target_app = str(data.get('target_app', '')).strip()
        except Exception:
            cfg.x = 0
            cfg.y = 0
            cfg.width = 320
            cfg.height = 180

        try:
            cfg.interval_ms = int(cfg.interval_ms or 2000)
        except Exception:
            cfg.interval_ms = 2000

        cfg.enabled = bool(cfg.enabled)
        if not cfg.name:
            cfg.name = 'Секція'

        # Міграція старих конфігів: якщо в JSON немає overlay_custom,
        # вважаємо, що секція тепер має власні налаштування оверлею.
        if 'overlay_custom' not in raw_keys:
            cfg.overlay_custom = True
        else:
            cfg.overlay_custom = bool(cfg.overlay_custom)

        # Санітизація пер-секційних полів
        if cfg.overlay_font_family is not None:
            s = str(cfg.overlay_font_family or '').strip()
            cfg.overlay_font_family = s if s else None

        if cfg.overlay_font_size is not None:
            try:
                cfg.overlay_font_size = float(cfg.overlay_font_size)
            except Exception:
                cfg.overlay_font_size = None

        if cfg.overlay_padding is not None:
            try:
                cfg.overlay_padding = int(cfg.overlay_padding)
            except Exception:
                cfg.overlay_padding = None

        if cfg.overlay_round_radius is not None:
            try:
                cfg.overlay_round_radius = int(cfg.overlay_round_radius)
            except Exception:
                cfg.overlay_round_radius = None

        if cfg.overlay_bg_color is not None:
            if isinstance(cfg.overlay_bg_color, (list, tuple)) and len(cfg.overlay_bg_color) == 3:
                try:
                    cfg.overlay_bg_color = [int(cfg.overlay_bg_color[0]), int(cfg.overlay_bg_color[1]), int(cfg.overlay_bg_color[2])]
                except Exception:
                    cfg.overlay_bg_color = None
            else:
                cfg.overlay_bg_color = None

        if cfg.overlay_bg_opacity is not None:
            try:
                cfg.overlay_bg_opacity = int(cfg.overlay_bg_opacity)
            except Exception:
                cfg.overlay_bg_opacity = None

        if cfg.overlay_text_color is not None:
            if isinstance(cfg.overlay_text_color, (list, tuple)) and len(cfg.overlay_text_color) == 3:
                try:
                    cfg.overlay_text_color = [int(cfg.overlay_text_color[0]), int(cfg.overlay_text_color[1]), int(cfg.overlay_text_color[2])]
                except Exception:
                    cfg.overlay_text_color = None
            else:
                cfg.overlay_text_color = None

        if cfg.overlay_text_opacity is not None:
            try:
                cfg.overlay_text_opacity = int(cfg.overlay_text_opacity)
            except Exception:
                cfg.overlay_text_opacity = None

        if cfg.overlay_use_ocr_bg_color is not None:
            cfg.overlay_use_ocr_bg_color = bool(cfg.overlay_use_ocr_bg_color)

        if cfg.overlay_max_chars_per_line is not None:
            try:
                cfg.overlay_max_chars_per_line = int(cfg.overlay_max_chars_per_line)
            except Exception:
                cfg.overlay_max_chars_per_line = None

        if cfg.overlay_max_box_height is not None:
            try:
                cfg.overlay_max_box_height = int(cfg.overlay_max_box_height)
            except Exception:
                cfg.overlay_max_box_height = None

        return cfg


def load_config() -> AppConfig:
    cfg = AppConfig()
    path = cfg.config_path()
    try:
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            roi_raw = None
            if isinstance(data, dict):
                roi_raw = data.get('roi_sections')

            for k, v in (data or {}).items():
                if k == 'roi_sections':
                    continue
                if hasattr(cfg, k):
                    setattr(cfg, k, v)

            # ROI секції (можуть бути відсутні у старих конфігах)
            roi_list: List[RoiSectionConfig] = []
            if isinstance(roi_raw, list):
                for row in roi_raw:
                    try:
                        roi_list.append(RoiSectionConfig.from_dict(row))
                    except Exception:
                        continue
            cfg.roi_sections = roi_list
    except Exception:
        # якщо файл битий, просто стартуємо з дефолтом
        pass
    return cfg


def save_config(cfg: AppConfig) -> None:
    """Зберігає конфіг атомарно і потокобезпечно.

    Причина: у нас є фонові потоки (ROI/continuous), які теж можуть
    викликати save_config. На Windows паралельні записи в один файл
    інколи дають PermissionError або залишають битий JSON.
    """

    os.makedirs(cfg.config_dir(), exist_ok=True)

    # Глобальний lock на весь модуль (щоб усі виклики йшли в один ряд).
    global _SAVE_LOCK

    path = cfg.config_path()
    tmp_path = path + '.tmp'

    payload = asdict(cfg)

    with _SAVE_LOCK:
        # Пишемо в тимчасовий файл, потім атомарно замінюємо.
        with open(tmp_path, 'w', encoding='utf-8') as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
            try:
                f.flush()
                os.fsync(f.fileno())
            except Exception:
                pass

        try:
            os.replace(tmp_path, path)
        except Exception:
            # fallback для випадків, коли replace/rename поводиться дивно
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass
            os.rename(tmp_path, path)
