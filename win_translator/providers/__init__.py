"""Провайдери OCR/перекладу + менеджер.

Це адаптація логіки з Decky-Translator під Windows-додаток.
"""

import logging
from typing import List, Optional

from .base import (
    OCRProvider,
    TranslationProvider,
    ProviderType,
    TextRegion,
    NetworkError,
    ApiKeyError,
    RateLimitError,
)
from .google_ocr import GoogleVisionProvider
from .google_translate import GoogleTranslateProvider
from .ocrspace import OCRSpaceProvider
from .free_translate import FreeTranslateProvider
from .rapidocr_provider import RapidOCRProvider
from .argos_translate import ArgosTranslateProvider

logger = logging.getLogger(__name__)

__all__ = [
    'OCRProvider',
    'TranslationProvider',
    'ProviderType',
    'TextRegion',
    'NetworkError',
    'ApiKeyError',
    'RateLimitError',
    'GoogleVisionProvider',
    'GoogleTranslateProvider',
    'OCRSpaceProvider',
    'FreeTranslateProvider',
    'RapidOCRProvider',
    'ArgosTranslateProvider',
    'ProviderManager',
]


class ProviderManager:
    """Factory/manager для OCR і перекладу."""

    def __init__(self, data_dir: str = "", ocrspace_api_key: str = "helloworld"):
        self._ocr_providers = {}
        self._translation_providers = {}

        # Windows-логіка за замовчуванням: без ключів.
        self._data_dir = data_dir or ""
        self._ocrspace_api_key = ocrspace_api_key or "helloworld"
        # OCR: OCR.space (хмара), переклад: безкоштовний Google endpoint.
        self._google_api_key = ""
        self._ocr_provider_preference = "ocrspace"  # rapidocr | ocrspace | googlecloud
        self._translation_provider_preference = "freegoogle"  # freegoogle | googlecloud | argos

        # RapidOCR тюнінг (опційно)
        self._rapidocr_confidence = 0.5
        self._rapidocr_box_thresh = 0.5
        self._rapidocr_unclip_ratio = 1.6

        # last used providers (for logs/debug)
        self.last_ocr_provider_name: str = ""
        self.last_translation_provider_name: str = ""

    def configure(
        self,
        google_api_key: str = "",
        ocr_provider: str = "",
        translation_provider: str = "",
        rapidocr_models_dir: str = "",
        ocrspace_api_key: str = "",
        data_dir: str = "",
    ) -> None:
        self._google_api_key = google_api_key or ""
        if ocrspace_api_key:
            self._ocrspace_api_key = ocrspace_api_key
        if data_dir:
            self._data_dir = data_dir
        if ocr_provider:
            self._ocr_provider_preference = ocr_provider
        if translation_provider:
            self._translation_provider_preference = translation_provider

        # Оновити ключі в існуючих інстансах
        if ProviderType.GOOGLE in self._ocr_providers:
            self._ocr_providers[ProviderType.GOOGLE].set_api_key(self._google_api_key)
        if ProviderType.GOOGLE in self._translation_providers:
            self._translation_providers[ProviderType.GOOGLE].set_api_key(self._google_api_key)

        # Якщо користувач передав шлях до моделей RapidOCR
        if rapidocr_models_dir:
            rapidocr = self._ocr_providers.get(ProviderType.RAPIDOCR)
            if rapidocr:
                rapidocr.set_models_dir(rapidocr_models_dir)

    def set_rapidocr_confidence(self, confidence: float) -> None:
        self._rapidocr_confidence = max(0.0, min(1.0, confidence))
        rapidocr = self._ocr_providers.get(ProviderType.RAPIDOCR)
        if rapidocr:
            rapidocr.set_min_confidence(self._rapidocr_confidence)

    def set_rapidocr_box_thresh(self, box_thresh: float) -> None:
        self._rapidocr_box_thresh = max(0.0, min(1.0, box_thresh))
        rapidocr = self._ocr_providers.get(ProviderType.RAPIDOCR)
        if rapidocr:
            rapidocr.set_box_thresh(self._rapidocr_box_thresh)

    def set_rapidocr_unclip_ratio(self, unclip_ratio: float) -> None:
        self._rapidocr_unclip_ratio = max(1.0, min(3.0, unclip_ratio))
        rapidocr = self._ocr_providers.get(ProviderType.RAPIDOCR)
        if rapidocr:
            rapidocr.set_unclip_ratio(self._rapidocr_unclip_ratio)

    def get_ocr_provider(self, provider_type: Optional[ProviderType] = None) -> Optional[OCRProvider]:
        if provider_type is None:
            if self._ocr_provider_preference == "rapidocr":
                provider_type = ProviderType.RAPIDOCR
            elif self._ocr_provider_preference == "googlecloud":
                provider_type = ProviderType.GOOGLE
            else:
                provider_type = ProviderType.OCR_SPACE

        if provider_type not in self._ocr_providers:
            if provider_type == ProviderType.OCR_SPACE:
                self._ocr_providers[provider_type] = OCRSpaceProvider(api_key=self._ocrspace_api_key, data_dir=self._data_dir)
            elif provider_type == ProviderType.GOOGLE:
                self._ocr_providers[provider_type] = GoogleVisionProvider(self._google_api_key)
            elif provider_type == ProviderType.RAPIDOCR:
                self._ocr_providers[provider_type] = RapidOCRProvider(
                    plugin_dir="",
                    min_confidence=self._rapidocr_confidence
                )
                # Підхопити тюнінги
                self._ocr_providers[provider_type].set_box_thresh(self._rapidocr_box_thresh)
                self._ocr_providers[provider_type].set_unclip_ratio(self._rapidocr_unclip_ratio)

        return self._ocr_providers.get(provider_type)

    def get_translation_provider(self, provider_type: Optional[ProviderType] = None) -> Optional[TranslationProvider]:
        if provider_type is None:
            pref = str(self._translation_provider_preference or '').strip().lower()
            if pref == 'googlecloud':
                provider_type = ProviderType.GOOGLE
            elif pref == 'argos':
                provider_type = ProviderType.ARGOS
            else:
                provider_type = ProviderType.FREE_GOOGLE

        if provider_type not in self._translation_providers:
            if provider_type == ProviderType.FREE_GOOGLE:
                self._translation_providers[provider_type] = FreeTranslateProvider()
            elif provider_type == ProviderType.GOOGLE:
                self._translation_providers[provider_type] = GoogleTranslateProvider(self._google_api_key)
            elif provider_type == ProviderType.ARGOS:
                self._translation_providers[provider_type] = ArgosTranslateProvider()


        return self._translation_providers.get(provider_type)

    async def recognize_text(self, image_data: bytes, language: str = "auto") -> List[TextRegion]:
        # Спочатку пробуємо вибраний провайдер
        provider = self.get_ocr_provider()
        if provider is not None:
            if provider.is_available(language):
                logger.debug(f"OCR: {provider.name}")
                self.last_ocr_provider_name = provider.name
                return await provider.recognize(image_data, language)

        # Якщо вибраний не доступний (частий кейс: RapidOCR без моделей),
        # робимо фолбек без додаткових питань.
        candidates: List[ProviderType] = [ProviderType.OCR_SPACE, ProviderType.GOOGLE, ProviderType.RAPIDOCR]

        for ptype in candidates:
            p = self.get_ocr_provider(ptype)
            if p is None:
                continue
            if not p.is_available(language):
                continue
            logger.debug(f"OCR fallback: {p.name}")
            self.last_ocr_provider_name = p.name
            return await p.recognize(image_data, language)

        logger.warning("Немає доступного OCR провайдера")
        self.last_ocr_provider_name = ""
        return []

    async def translate_text(self, texts: List[str], source_lang: str, target_lang: str) -> List[str]:
        if not texts:
            return []

        pref = str(self._translation_provider_preference or '').strip().lower()

        # Будуємо список кандидатів у порядку пріоритету.
        candidates: List[ProviderType] = []
        if pref == 'argos':
            candidates = [ProviderType.ARGOS, ProviderType.GOOGLE, ProviderType.FREE_GOOGLE]
        elif pref == 'googlecloud':
            candidates = [ProviderType.GOOGLE, ProviderType.FREE_GOOGLE, ProviderType.ARGOS]
        else:
            candidates = [ProviderType.FREE_GOOGLE, ProviderType.GOOGLE, ProviderType.ARGOS]

        tried: List[str] = []
        for ptype in candidates:
            provider = self.get_translation_provider(ptype)
            if provider is None:
                continue

            tried.append(str(getattr(provider, 'name', '') or str(ptype.value)))

            ok = False
            try:
                ok = bool(provider.is_available(source_lang, target_lang))
            except Exception:
                ok = False

            if not ok:
                continue

            self.last_translation_provider_name = provider.name

            # Логи для перевірки Argos
            try:
                if provider.provider_type == ProviderType.ARGOS:
                    logger.info('Translate: using ARGOS (%s->%s) items=%d', source_lang, target_lang, len(texts))
                else:
                    logger.debug('Translate: %s', provider.name)
            except Exception:
                pass

            return await provider.translate_batch(texts, source_lang, target_lang)

        logger.warning('Немає доступного провайдера перекладу (%s->%s). tried=%s', source_lang, target_lang, tried)
        self.last_translation_provider_name = ''
        return texts
        return texts
