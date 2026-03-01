import asyncio
import logging
from typing import List, Optional

from .base import TranslationProvider, ProviderType


logger = logging.getLogger(__name__)


def _norm_lang(code: str) -> str:
    s = str(code or '').strip().lower()
    if not s:
        return ''
    # Argos не вміє авто-детект.
    if s == 'auto':
        return ''
    # Часто прилітає щось типу en-US.
    if '-' in s:
        s = s.split('-', 1)[0]
    if '_' in s:
        s = s.split('_', 1)[0]
    return s


class ArgosTranslateProvider(TranslationProvider):
    """Локальний переклад через Argos Translate.

    Це опційний провайдер: якщо бібліотека не встановлена або немає
    інстальованого мовного пакета для пари мов, він вважається недоступним,
    і ProviderManager спокійно зробить фолбек на інший переклад.
    """

    def __init__(self):
        self._ok = False
        self._warned_pairs = set()
        self._translate = None
        self._installed_langs = None
        try:
            import argostranslate.translate as _tr
            self._translate = _tr
            self._installed_langs = _tr.get_installed_languages
            self._ok = True
            logger.debug('ArgosTranslateProvider: library loaded')
        except Exception as e:
            self._ok = False
            logger.debug('ArgosTranslateProvider: not available (%s)', e)

    @property
    def name(self) -> str:
        return 'argos'

    @property
    def provider_type(self) -> ProviderType:
        return ProviderType.ARGOS

    def _get_translation(self, source_lang: str, target_lang: str):
        if not self._ok:
            return None
        src = _norm_lang(source_lang)
        tgt = _norm_lang(target_lang)
        if not src or not tgt:
            return None

        try:
            langs = self._installed_langs()
        except Exception:
            return None

        src_obj = None
        tgt_obj = None
        for l in (langs or []):
            try:
                if str(getattr(l, 'code', '') or '').lower() == src:
                    src_obj = l
                if str(getattr(l, 'code', '') or '').lower() == tgt:
                    tgt_obj = l
            except Exception:
                continue

        if src_obj is None or tgt_obj is None:
            return None

        try:
            return src_obj.get_translation(tgt_obj)
        except Exception:
            return None

    def is_available(self, source_lang: str, target_lang: str) -> bool:
        if not self._ok:
            return False
        tr = self._get_translation(source_lang, target_lang)
        return tr is not None

    async def translate(self, text: str, source_lang: str, target_lang: str) -> str:
        out_list = await self.translate_batch([text], source_lang, target_lang)
        if not out_list:
            return ''
        return str(out_list[0] or '')

    async def translate_batch(self, texts: List[str], source_lang: str, target_lang: str) -> List[str]:
        if not texts:
            return []

        tr = self._get_translation(source_lang, target_lang)
        if tr is None:
            # Немає пакета, віддаємо як є (але ProviderManager зазвичай сюди не зайде,
            # бо is_available() поверне False). Це «на всяк випадок».
            try:
                key = f"{_norm_lang(source_lang)}->{_norm_lang(target_lang)}"
                if key not in self._warned_pairs:
                    self._warned_pairs.add(key)
                    logger.warning('Argos: no installed language package for %s', key)
            except Exception:
                pass
            return list(texts)

        loop = None
        try:
            loop = asyncio.get_running_loop()
        except Exception:
            loop = None

        def _do_translate(batch: List[str]) -> List[str]:
            out: List[str] = []
            for t in (batch or []):
                try:
                    out.append(str(tr.translate(str(t or ''))))
                except Exception:
                    out.append(str(t or ''))
            return out

        # Переклад синхронний, тому віддаємо в threadpool, щоб не блокувати event loop.
        if loop is None:
            return _do_translate(texts)

        try:
            return await loop.run_in_executor(None, _do_translate, list(texts))
        except Exception:
            logger.exception('Argos translate failed')
            return list(texts)

    def get_supported_languages(self) -> List[str]:
        if not self._ok:
            return []
        try:
            langs = self._installed_langs()
            out: List[str] = []
            for l in (langs or []):
                try:
                    code = str(getattr(l, 'code', '') or '').strip().lower()
                    if code:
                        out.append(code)
                except Exception:
                    continue
            return out
        except Exception:
            return []
