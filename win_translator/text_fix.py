import logging
import re
from typing import List, Optional

logger = logging.getLogger(__name__)


_WORD_RE = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ]+")


_spell_cache = {}


def _get_spellchecker(lang: str):
    code = str(lang or '').strip().lower()
    if not code:
        return None

    # pyspellchecker: 'en', 'de', 'es', 'fr', 'pt', 'ru', ... (але найстабільніше en)
    if '-' in code:
        code = code.split('-', 1)[0]
    if '_' in code:
        code = code.split('_', 1)[0]

    if code in _spell_cache:
        return _spell_cache.get(code)

    try:
        from spellchecker import SpellChecker  # pyspellchecker
        sp = SpellChecker(language=code)
        _spell_cache[code] = sp
        logger.info('Spellcheck: pyspellchecker loaded (lang=%s)', code)
        return sp
    except Exception as e:
        _spell_cache[code] = None
        logger.debug('Spellcheck: pyspellchecker not available for lang=%s (%s)', code, e)
        return None


def _should_skip_word(word: str) -> bool:
    w = str(word or '')
    if not w:
        return True

    # Короткі слова частіше ламаються, ніж лікуються
    if len(w) < 3:
        return True

    # Цифри / змішані токени
    for ch in w:
        if ch.isdigit():
            return True

    # Апострофи/дефіси: у OCR часто це коректні конструкції (shan't, don't, re-enter)
    if "'" in w or "’" in w or "-" in w:
        return True

    # ALLCAPS схоже на абревіатури
    if w.isupper() and len(w) <= 6:
        return True

    # Назви/імена (TitleCase): краще не чіпати
    if w[:1].isupper():
        return True

    return False


def fix_text(text: str, lang: str = 'en') -> str:
    """Легка спроба виправити OCR-помилки перед перекладом.

    Робить мінімально агресивні заміни, щоб не псувати імена/терміни.
    Якщо бібліотек немає, повертає оригінал.
    """

    src = str(text or '')
    if not src.strip():
        return src

    sp = _get_spellchecker(lang)
    if sp is None:
        return src

    # Працюємо по словах, зберігаючи все інше як є.
    out = []
    last = 0

    for m in _WORD_RE.finditer(src):
        start, end = m.span()
        out.append(src[last:start])

        word = m.group(0)
        if _should_skip_word(word):
            out.append(word)
            last = end
            continue

        w_low = word.lower()
        try:
            if w_low in sp:
                out.append(word)
                last = end
                continue
        except Exception:
            # якщо SpellChecker капризує, просто не чіпаємо
            out.append(word)
            last = end
            continue

        try:
            cand = sp.correction(w_low)
        except Exception:
            cand = None

        if not cand:
            out.append(word)
            last = end
            continue

        cand = str(cand)
        if cand == w_low:
            out.append(word)
            last = end
            continue

        # Мінімальні запобіжники проти дивних замін
        if cand[:1] != w_low[:1]:
            out.append(word)
            last = end
            continue

        if abs(len(cand) - len(w_low)) > 1:
            out.append(word)
            last = end
            continue

        out.append(cand)
        last = end

    out.append(src[last:])
    return ''.join(out)


def fix_text_batch(texts: List[str], lang: str = 'en') -> List[str]:
    if not texts:
        return []

    out: List[str] = []
    for t in (texts or []):
        out.append(fix_text(t, lang=lang))
    return out
