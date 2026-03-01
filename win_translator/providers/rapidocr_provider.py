# providers/rapidocr_provider.py
# Local RapidOCR provider - runs entirely on device without internet
# Windows build: NO subprocess. Everything runs in-process.

import logging
import os
import time
import numbers
from typing import List, Optional, Dict, Any

from .base import OCRProvider, ProviderType, TextRegion

logger = logging.getLogger(__name__)

DEFAULT_MIN_CONFIDENCE = 0.5
RAPIDOCR_MODELS_DIR = "models/rapidocr"

# Maximum image dimension for OCR (resize larger images for performance)
# 1920 часто занадто агресивно для дрібного тексту на 2К/4К.
MAX_IMAGE_DIMENSION = 3200


class RapidOCRProvider(OCRProvider):
    """
    OCR provider using RapidOCR locally (no subprocess).

    Strategy:
      - Prefer `rapidocr` package (RapidAI/RapidOCR) for simplicity.
      - If custom model paths are present OR `rapidocr` import fails,
        fallback to `rapidocr_onnxruntime` which supports explicit det/rec/cls paths.

    Notes:
      - On Windows with multiple monitors and scaling, OCR coords should be mapped
        by the caller (overlay layer). Provider returns pixel coords in the input image space.
    """

    LANGUAGE_MAP = {
        "auto": "ch",
        "en": "en",
        "zh-CN": "ch",
        "zh-TW": "ch",
        "ja": "ch",
        "ko": "korean",
        "de": "latin",
        "fr": "latin",
        "es": "latin",
        "it": "latin",
        "pt": "latin",
        "nl": "latin",
        "pl": "latin",
        "tr": "latin",
        "ro": "latin",
        "vi": "latin",
        "fi": "latin",
        "ru": "eslav",
        "uk": "eslav",
        "bg": "eslav",
        "el": "greek",
        "th": "thai",
    }

    SUPPORTED_LANGUAGES = [
        "auto", "en", "zh-CN", "zh-TW", "ja", "ko",
        "de", "fr", "es", "it", "pt", "nl", "pl", "tr", "ro", "vi", "fi",
        "ru", "uk", "bg", "el", "th"
    ]

    def __init__(
        self,
        plugin_dir: str = "",
        min_confidence: float = DEFAULT_MIN_CONFIDENCE
    ):
        default_root = os.environ.get("WIN_TRANSLATOR_HOME")
        if not default_root:
            default_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

        self._plugin_dir = plugin_dir or default_root
        self._models_dir = os.path.join(self._plugin_dir, RAPIDOCR_MODELS_DIR)

        self._min_confidence = max(0.0, min(1.0, min_confidence))
        self._box_thresh = 0.5
        self._unclip_ratio = 1.6

        self._available: Optional[bool] = None
        self._init_error: Optional[str] = None

        # Engines cached per "lang family"
        self._engine_cache: Dict[str, Any] = {}

        logger.debug(
            "RapidOCRProvider initialized (plugin_dir=%s, models_dir=%s, min_confidence=%s)",
            self._plugin_dir,
            self._models_dir,
            self._min_confidence
        )

    @property
    def name(self) -> str:
        return "RapidOCR (Local)"

    @property
    def provider_type(self) -> ProviderType:
        return ProviderType.RAPIDOCR

    def get_supported_languages(self) -> List[str]:
        return self.SUPPORTED_LANGUAGES.copy()

    def set_models_dir(self, models_dir: str) -> None:
        if models_dir:
            self._models_dir = models_dir
            self._available = None
            self._engine_cache = {}

    def set_min_confidence(self, confidence: float) -> None:
        self._min_confidence = max(0.0, min(1.0, confidence))

    def set_box_thresh(self, box_thresh: float) -> None:
        self._box_thresh = max(0.0, min(1.0, box_thresh))

    def set_unclip_ratio(self, unclip_ratio: float) -> None:
        self._unclip_ratio = max(1.0, min(3.0, unclip_ratio))

    def get_init_error(self) -> Optional[str]:
        return self._init_error

    def _has_custom_models(self) -> bool:
        # Мінімальна перевірка, що в папці є хоча б det+cls.
        # Підтримуємо і v5 (mobile_det) і v4 (det_infer), бо люди люблять хаос.
        cls_model = os.path.join(self._models_dir, "ch_ppocr_mobile_v2.0_cls_infer.onnx")
        det_v5 = os.path.join(self._models_dir, "ch_PP-OCRv5_mobile_det.onnx")
        det_v4 = os.path.join(self._models_dir, "ch_PP-OCRv4_det_infer.onnx")

        if os.path.exists(cls_model) and (os.path.exists(det_v5) or os.path.exists(det_v4)):
            return True
        return False

    def _check_availability(self) -> bool:
        self._init_error = None

        # 1) Спроба імпорту rapidocr
        try:
            import rapidocr  # noqa: F401
            return True
        except Exception as e:
            logger.info("rapidocr import failed: %s", e)

        # 2) Якщо rapidocr нема, пробуємо rapidocr_onnxruntime
        try:
            import rapidocr_onnxruntime  # noqa: F401
            return True
        except Exception as e:
            self._init_error = "Не знайдено ні rapidocr, ні rapidocr_onnxruntime"
            logger.warning("%s (%s)", self._init_error, e)
            return False

    def is_available(self, language: str = "auto") -> bool:
        if self._available is None:
            self._available = self._check_availability()

        if not self._available:
            return False

        if language not in self.SUPPORTED_LANGUAGES:
            return False

        return True

    def _build_engine_rapidocr(self) -> Any:
        # rapidocr: просте використання без явних шляхів моделей
        from rapidocr import RapidOCR
        engine = RapidOCR()
        return engine

    def _pick_rec_model_for_family(self, family: str) -> Optional[str]:
        # Ти можеш назвати як завгодно, але даю дефолтні очікувані імена
        candidates = []
        candidates.append(os.path.join(self._models_dir, f"{family}_rec.onnx"))
        candidates.append(os.path.join(self._models_dir, f"{family}_PP-OCRv5_rec.onnx"))
        candidates.append(os.path.join(self._models_dir, "rec.onnx"))

        for p in candidates:
            if os.path.exists(p):
                return p

        return None

    def _pick_rec_keys_for_family(self, family: str) -> Optional[str]:
        candidates = []
        candidates.append(os.path.join(self._models_dir, f"{family}_dict.txt"))
        candidates.append(os.path.join(self._models_dir, f"{family}_keys.txt"))
        candidates.append(os.path.join(self._models_dir, "dict.txt"))
        candidates.append(os.path.join(self._models_dir, "keys.txt"))

        for p in candidates:
            if os.path.exists(p):
                return p

        return None

    def _build_engine_onnxruntime(self, family: str) -> Any:
        # rapidocr_onnxruntime дає явні шляхи моделей (det/rec/cls)
        # Це найзручніший спосіб підключити свої ONNX.
        from rapidocr_onnxruntime import RapidOCR

        # det: v5 preferred, fallback v4
        det_model = os.path.join(self._models_dir, "ch_PP-OCRv5_mobile_det.onnx")
        if not os.path.exists(det_model):
            det_model = os.path.join(self._models_dir, "ch_PP-OCRv4_det_infer.onnx")

        cls_model = os.path.join(self._models_dir, "ch_ppocr_mobile_v2.0_cls_infer.onnx")
        rec_model = self._pick_rec_model_for_family(family)

        # rec_keys можуть бути потрібні залежно від моделі
        # (деякі збірки rapidocr_onnxruntime беруть ключі з пакета,
        # але якщо ти використовуєш свої, краще дати явно)
        rec_keys = self._pick_rec_keys_for_family(family)

        kwargs = {
            "det_model_path": det_model,
            "cls_model_path": cls_model,
        }

        # Якщо немає family-специфічного rec, беремо дефолтні ch-моделі (v4/v5)
        if not rec_model:
            rec_model = os.path.join(self._models_dir, "ch_PP-OCRv5_rec_infer.onnx")
            if not os.path.exists(rec_model):
                rec_model = os.path.join(self._models_dir, "ch_PP-OCRv4_rec_infer.onnx")

        if rec_model and os.path.exists(rec_model):
            kwargs["rec_model_path"] = rec_model

        if not rec_keys:
            # найпоширеніший файл ключів у PaddleOCR
            cand = os.path.join(self._models_dir, "ppocr_keys_v1.txt")
            if os.path.exists(cand):
                rec_keys = cand

        if rec_keys and os.path.exists(rec_keys):
            kwargs["rec_keys_path"] = rec_keys

        # Пороги детекції: тримай консервативно, якщо не знаєш
        # (в rapidocr_onnxruntime назви параметрів можуть різнитись по версіях,
        # тому тут не чіпаю det_thresh/limit_side_len, щоб не впасти)
        engine = RapidOCR(**kwargs)
        return engine

    def _get_engine(self, family: str) -> Any:
        if family in self._engine_cache:
            return self._engine_cache[family]

        # Якщо є кастомні моделі, краще використати onnxruntime-движок з явними шляхами.
        custom_models = self._has_custom_models()

        if not custom_models:
            # Спробуємо rapidocr (найпростіше)
            try:
                engine = self._build_engine_rapidocr()
                self._engine_cache[family] = engine
                return engine
            except Exception as e:
                logger.info("rapidocr engine init failed, fallback to onnxruntime: %s", e)

        # Fallback: rapidocr_onnxruntime
        engine = self._build_engine_onnxruntime(family)
        self._engine_cache[family] = engine
        return engine

    def _decode_image_to_bgr(self, image_data: bytes):
        # Не тягнемо cv2 як обовʼязковий декодер: PIL стабільніший для bytes.
        import io
        from PIL import Image
        import numpy as np

        img = Image.open(io.BytesIO(image_data))
        img = img.convert("RGB")
        arr = np.array(img)  # RGB
        # Convert RGB -> BGR for OpenCV-style pipelines if needed
        bgr = arr[:, :, ::-1].copy()
        return bgr

    def _ensure_max_dimension(self, bgr):
        import cv2

        h, w = bgr.shape[0], bgr.shape[1]
        max_dim = w
        if h > max_dim:
            max_dim = h

        if max_dim <= MAX_IMAGE_DIMENSION:
            return bgr, 1.0, 1.0

        scale = float(MAX_IMAGE_DIMENSION) / float(max_dim)
        new_w = int(w * scale)
        new_h = int(h * scale)

        if new_w < 1:
            new_w = 1
        if new_h < 1:
            new_h = 1

        resized = cv2.resize(bgr, (new_w, new_h), interpolation=cv2.INTER_AREA)
        sx = float(w) / float(new_w)
        sy = float(h) / float(new_h)
        return resized, sx, sy

    async def recognize(self, image_data: bytes, language: str = "auto") -> List[TextRegion]:
        if self._available is None:
            self._available = self._check_availability()

        if not self._available:
            return []

        if language not in self.SUPPORTED_LANGUAGES:
            return []

        # Lazy imports
        import io
        from PIL import Image
        import numpy as np

        start = time.time()

        # Decode bytes -> numpy image
        try:
            img = Image.open(io.BytesIO(image_data))
            img = img.convert("RGB")
            rgb = np.array(img)
        except Exception as e:
            logger.error("RapidOCR: failed to decode image bytes: %s", e)
            return []

        # Resize for speed if needed
        try:
            import cv2
            bgr = rgb[:, :, ::-1].copy()
            bgr_small, sx, sy = self._ensure_max_dimension(bgr)
        except Exception as e:
            logger.error("RapidOCR: resize failed: %s", e)
            return []

        family = self.LANGUAGE_MAP.get(language, "ch")
        engine = None
        try:
            print(f"RapidOCR: using language family '{family}' for language '{language}'")
            engine = self._get_engine(family)
        except Exception as e:
            logger.error("RapidOCR: engine init failed: %s", e)
            return []

        # Run OCR
        try:
            # rapidocr: зазвичай повертає список (як PaddleOCR)
            # rapidocr_onnxruntime: часто повертає (result, elapse_list)
            out = engine(bgr_small)
            raw_result = out[0] if isinstance(out, tuple) and len(out) > 0 else out

            regions = self._parse_result(raw_result, sx, sy)

            print(f"RapidOCR: found {regions}")

            # Деякі збірки очікують RGB ndarray, а не BGR.
            # Якщо нічого не знайшли, пробуємо ще раз з RGB.
            if not regions:
                try:
                    rgb_small = bgr_small[:, :, ::-1].copy()
                    out2 = engine(rgb_small)
                    raw2 = out2[0] if isinstance(out2, tuple) and len(out2) > 0 else out2
                    regions2 = self._parse_result(raw2, sx, sy)
                    if len(regions2) > len(regions):
                        regions = regions2
                except Exception:
                    pass

            elapsed = time.time() - start
            logger.debug("RapidOCR: %d regions in %.2fs", len(regions), elapsed)
            return regions

        except Exception as e:
            logger.error("RapidOCR: OCR error: %s", e, exc_info=True)
            return []

    def _parse_result(self, raw_result, sx: float, sy: float) -> List[TextRegion]:
        regions: List[TextRegion] = []

        if raw_result is None:
            return regions

        # ✅ Shape 0: RapidOCROutput (object with attributes)
        # rapidocr returns RapidOCROutput(boxes=..., txts=..., scores=...)
        if hasattr(raw_result, "boxes") and hasattr(raw_result, "txts"):
            try:
                boxes = raw_result.boxes
                txts = raw_result.txts
                scores = getattr(raw_result, "scores", None)

                # numpy -> list
                if hasattr(boxes, "tolist"):
                    boxes = boxes.tolist()

                # txts can be tuple[str]
                if isinstance(txts, tuple):
                    txts = list(txts)

                # scores can be tuple[float] or np
                if scores is not None and hasattr(scores, "tolist"):
                    scores = scores.tolist()
                if isinstance(scores, tuple):
                    scores = list(scores)

                if not isinstance(boxes, list) or not isinstance(txts, list):
                    return regions

                n = len(txts)
                for i in range(n):
                    text = txts[i]
                    box = boxes[i] if i < len(boxes) else None
                    score = None
                    if isinstance(scores, list) and i < len(scores):
                        score = scores[i]

                    region = self._build_region_from_parts(box, text, score, sx, sy)
                    if region:
                        regions.append(region)

                return regions

            except Exception as e:
                logger.error("RapidOCR: failed to parse RapidOCROutput: %s", e, exc_info=True)
                return regions

        # Shape A: list/tuple of entries
        if isinstance(raw_result, (list, tuple)):
            for item in raw_result:
                region = self._parse_single_item(item, sx, sy)
                if region:
                    regions.append(region)
            return regions

        # Shape B: dict-ish
        if isinstance(raw_result, dict):
            boxes = raw_result.get("boxes")
            texts = raw_result.get("texts")
            scores = raw_result.get("scores")

            if isinstance(boxes, list) and isinstance(texts, list):
                for i in range(len(texts)):
                    text = texts[i]
                    score = None
                    if isinstance(scores, list) and i < len(scores):
                        score = scores[i]
                    box = None
                    if i < len(boxes):
                        box = boxes[i]

                    region = self._build_region_from_parts(box, text, score, sx, sy)
                    if region:
                        regions.append(region)

            return regions

        # Unknown format
        try:
            logger.debug("RapidOCR: unknown result format: %s", type(raw_result))
        except Exception:
            pass
        return regions

    def _parse_single_item(self, item, sx: float, sy: float) -> Optional[TextRegion]:
        if not isinstance(item, (list, tuple)):
            return None

        if len(item) < 2:
            return None

        box = item[0]

        text = None
        score = None

        second = item[1]

        # Випадок: [box, "text", score]
        if isinstance(second, str):
            text = second
            if len(item) >= 3 and isinstance(item[2], numbers.Real):
                score = float(item[2])

        # Випадок: [box, ("text", score)] або [box, ["text", score]]
        if text is None and isinstance(second, (list, tuple)) and len(second) >= 1:
            if isinstance(second[0], str):
                text = second[0]
                if len(second) >= 2 and isinstance(second[1], numbers.Real):
                    score = float(second[1])

        if text is None:
            return None

        return self._build_region_from_parts(box, text, score, sx, sy)


    def _build_region_from_parts(self, box, text, score, sx: float, sy: float) -> Optional[TextRegion]:
        if not isinstance(text, str):
            return None

        conf = 1.0
        if isinstance(score, numbers.Real):
            conf = float(score)

        if conf < self._min_confidence:
            return None

        rect = self._box_to_rect(box, sx, sy)
        if rect is None:
            return None

        return TextRegion(
            text=text,
            rect=rect,              # <-- ВАЖЛИВО: dict, не tuple
            confidence=conf,
            is_dialog=False
        )

    def _box_to_rect(self, box, sx: float, sy: float):
        # rapidocr / rapidocr_onnxruntime інколи повертають numpy.ndarray.
        # Нам потрібні звичайні списки.
        if box is not None and hasattr(box, "tolist"):
            try:
                box = box.tolist()
            except Exception:
                pass

        if box is None:
            return None

        pts = None

        # [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
        if isinstance(box, list) and len(box) == 4:
            ok = True
            fixed = []
            for p in box:
                if p is None:
                    ok = False
                    break
                if hasattr(p, "tolist"):
                    try:
                        p = p.tolist()
                    except Exception:
                        pass
                if not isinstance(p, (list, tuple)) or len(p) < 2:
                    ok = False
                    break
                fixed.append([p[0], p[1]])

            if ok:
                pts = fixed

        # [x1,y1,x2,y2,x3,y3,x4,y4]
        if pts is None and isinstance(box, (list, tuple)) and len(box) >= 8:
            pts = [
                [box[0], box[1]],
                [box[2], box[3]],
                [box[4], box[5]],
                [box[6], box[7]],
            ]

        if pts is None:
            return None

        xs = [float(p[0]) for p in pts]
        ys = [float(p[1]) for p in pts]

        min_x = min(xs) * sx
        max_x = max(xs) * sx
        min_y = min(ys) * sy
        max_y = max(ys) * sy

        left = int(round(min_x))
        top = int(round(min_y))
        right = int(round(max_x))
        bottom = int(round(max_y))

        if right <= left or bottom <= top:
            return None

        return {
            "left": left,
            "top": top,
            "right": right,
            "bottom": bottom,
        }

