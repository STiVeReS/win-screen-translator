import asyncio
import logging
from typing import List

# ВИПРАВЛЕНО: Імпортуємо TextRegion замість OcrRegion
from .base import TextRegion

logger = logging.getLogger(__name__)

try:
    from winsdk.windows.media.ocr import OcrEngine
    from winsdk.windows.globalization import Language
    from winsdk.windows.graphics.imaging import BitmapDecoder
    from winsdk.windows.storage.streams import DataWriter, InMemoryRandomAccessStream
    WIN_OCR_AVAILABLE = True
except ImportError:
    WIN_OCR_AVAILABLE = False
    logger.warning("winsdk не встановлено. Windows OCR недоступний.")


class WindowsOcrProvider:
    def __init__(self):
        self.is_available = WIN_OCR_AVAILABLE

    # ВИПРАВЛЕНО: Повертаємо List[TextRegion]
    async def recognize(self, image_bytes: bytes, lang: str) -> List[TextRegion]:
        if not self.is_available:
            logger.error("winsdk не встановлено. Виконайте 'pip install winsdk'")
            return []

        # Конвертуємо стандартні коди мов у формат Windows
        win_lang_code = lang
        if len(lang) == 2:
            mapping = {
                'en': 'en-US', 
                'uk': 'uk-UA', 
                'ru': 'ru-RU', 
                'ja': 'ja-JP',
                'zh': 'zh-Hans-CN',
                'de': 'de-DE',
                'fr': 'fr-FR',
                'es': 'es-ES'
            }
            win_lang_code = mapping.get(lang.lower(), lang)

        language = Language(win_lang_code)
        
        # Перевіряємо, чи встановлений мовний пакет у Windows
        if not OcrEngine.is_language_supported(language):
            logger.warning(f"Мова {win_lang_code} не підтримується Windows OCR. Спроба використати системну.")
            engine = OcrEngine.try_create_from_user_profile_languages()
        else:
            engine = OcrEngine.try_create_from_language(language)

        if engine is None:
            logger.error("Не вдалося ініціалізувати Windows OCR Engine.")
            return []

        # Завантажуємо байти зображення у потік для Windows API
        stream = InMemoryRandomAccessStream()
        writer = DataWriter(stream)
        writer.write_bytes(list(image_bytes))
        await writer.store_async()
        await writer.flush_async()
        stream.seek(0)

        # Декодуємо зображення
        decoder = await BitmapDecoder.create_async(stream)
        bitmap = await decoder.get_software_bitmap_async()

        # Запускаємо розпізнавання
        result = await engine.recognize_async(bitmap)

        regions: List[TextRegion] = []
        if result and result.lines:
            for line in result.lines:
                rect = line.bounding_rect
                # ВИПРАВЛЕНО: Створюємо об'єкт TextRegion
                regions.append(
                    TextRegion(
                        text=line.text,
                        rect={
                            'left': int(rect.x),
                            'top': int(rect.y),
                            'width': int(rect.width),
                            'height': int(rect.height),
                            'right': int(rect.x + rect.width),
                            'bottom': int(rect.y + rect.height),
                        }
                    )
                )
        return regions