import logging
import re
from typing import List, Optional
import pkg_resources

try:
    from symspellpy import SymSpell, Verbosity
except ImportError:
    pass

logger = logging.getLogger(__name__)

# Регулярний вираз для слів (включає апострофи, щоб don't не розбивалось на don і t)
_WORD_RE = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ\']+")

_sym_cache = {}
_spell_cache = {}


def _get_symspell(lang: str = 'en'):
    if lang not in _sym_cache:
        try:
            # Максимальна дистанція редагування 2, довжина префіксу 7
            sym_spell = SymSpell(max_dictionary_edit_distance=2, prefix_length=7)
            
            if lang == 'en':
                # Завантажуємо англійський словник, який іде з бібліотекою
                dictionary_path = pkg_resources.resource_filename(
                    "symspellpy", "frequency_dictionary_en_82_765.txt"
                )
                # Завантажуємо біграми для розуміння контексту сусідніх слів
                bigram_path = pkg_resources.resource_filename(
                    "symspellpy", "frequency_bigramdictionary_en_243_342.txt"
                )
                sym_spell.load_dictionary(dictionary_path, term_index=0, count_index=1)
                sym_spell.load_bigram_dictionary(bigram_path, term_index=0, count_index=2)
                _sym_cache[lang] = sym_spell
                logger.info('Spellcheck: SymSpell loaded for lang=%s', lang)
            else:
                # Для інших мов треба шукати частотні словники (поки вимикаємо)
                _sym_cache[lang] = None 
        except Exception as e:
            logger.debug('Spellcheck: SymSpell failed to load for lang=%s (%s)', lang, e)
            _sym_cache[lang] = None
            
    return _sym_cache[lang]


def _get_spellchecker(lang: str):
    code = str(lang or '').strip().lower()
    if not code:
        return None

    # pyspellchecker: 'en', 'de', 'es', 'fr', 'pt', 'ru', ... 
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

    # Менше 3 символів краще не чіпати (I, do, on, in і т.д.)
    if len(w) < 3:
        return True

    # Цифри / змішані токени
    for ch in w:
        if ch.isdigit():
            return True

    # Якщо слово починається чи закінчується на апостроф/дефіс - пропускаємо
    if w.startswith("'") or w.endswith("'") or w.startswith("-") or w.endswith("-"):
        return True

    # ALLCAPS - це найчастіше абревіатури
    if w.isupper():
        return True

    # TitleCase або camelCase імена залишаємо як є
    if w[:1].isupper() and any(c.islower() for c in w[1:]):
        return True

    return False


def fix_text(text: str, lang: str = 'en') -> str:
    """Легка спроба виправити OCR-помилки перед перекладом зі збереженням пунктуації."""
    src = str(text or '')
    if not src.strip():
        return src

    code = str(lang or 'en').strip().lower()
    if '-' in code:
        code = code.split('-', 1)[0]
    if '_' in code:
        code = code.split('_', 1)[0]

    # Визначаємо, який рушій використовувати
    use_symspell = False
    sp = None
    sym = None
    
    if code == 'en':
        sym = _get_symspell(code)
        if sym:
            use_symspell = True
            
    if not use_symspell:
        sp = _get_spellchecker(code)
        if sp is None:
            return src

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
        cand = None

        if use_symspell:
            try:
                # SymSpell пошук близького слова (Verbosity.CLOSEST)
                suggestions = sym.lookup(w_low, Verbosity.CLOSEST, max_edit_distance=2)
                if suggestions:
                    cand = suggestions[0].term
            except Exception:
                cand = None
        else:
            try:
                # Fallback до pyspellchecker для інших мов
                if w_low in sp or w_low.replace("'", "") in sp:
                    out.append(word)
                    last = end
                    continue
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

        # Мінімальні запобіжники проти дивних замін (дозволяємо різницю у 2 символи)
        if abs(len(cand) - len(w_low)) > 2:  
            out.append(word)
            last = end
            continue
            
        # Зберігаємо регістр оригінального слова
        if word.istitle():
            cand = cand.capitalize()
        elif word.isupper():
            cand = cand.upper()

        out.append(cand)
        last = end

    out.append(src[last:])
    return ''.join(out)


def fix_text_compound(text: str, lang: str = 'en') -> str:
    """
    Альтернативний метод: виправляє все речення разом (злиті або розірвані слова).
    Увага: Цей метод ідеальний для речень з помилками пробілів, але може втрачати оригінальну пунктуацію.
    """
    src = str(text or '')
    if not src.strip():
        return src

    code = str(lang or 'en').strip().lower()
    if '-' in code:
        code = code.split('-', 1)[0]
    if '_' in code:
        code = code.split('_', 1)[0]

    if code == 'en':
        sym = _get_symspell(code)
        if sym:
            try:
                suggestions = sym.lookup_compound(src, max_edit_distance=2)
                if suggestions:
                    corrected = suggestions[0].term
                    
                    # Легке відновлення великої літери на початку речення, якщо вона там була
                    if src and src[0].isupper() and corrected:
                        corrected = corrected[0].upper() + corrected[1:]
                        
                    return corrected
            except Exception:
                pass

    # Якщо мова не 'en' або SymSpell впав — повертаємося до стандартного методу
    return fix_text(src, lang)


def fix_text_batch(texts: List[str], lang: str = 'en') -> List[str]:
    if not texts:
        return []

    out: List[str] = []
    for t in (texts or []):
        # За замовчуванням використовуємо fix_text (зберігає пунктуацію).
        # Якщо хочете агресивніше виправляти пробіли ціною крапок і ком - змініть на fix_text_compound(t, lang=lang)
        out.append(fix_text(t, lang=lang))
    return out