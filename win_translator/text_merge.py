from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

from .providers.base import TextRegion


def _safe_float(x: object, default: float = 0.0) -> float:
    try:
        return float(x)
    except Exception:
        return float(default)


def _rect_to_ltrb(rect: object) -> Optional[Tuple[float, float, float, float]]:
    """Нормалізує rect у (left, top, right, bottom)."""
    if rect is None:
        return None

    if isinstance(rect, (list, tuple)) and len(rect) == 4:
        x = _safe_float(rect[0], 0.0)
        y = _safe_float(rect[1], 0.0)
        w = _safe_float(rect[2], 0.0)
        h = _safe_float(rect[3], 0.0)
        return (x, y, x + w, y + h)

    if isinstance(rect, dict):
        left = _safe_float(rect.get('left', 0.0), 0.0)
        top = _safe_float(rect.get('top', 0.0), 0.0)

        if rect.get('right') is not None and rect.get('bottom') is not None:
            right = _safe_float(rect.get('right', 0.0), 0.0)
            bottom = _safe_float(rect.get('bottom', 0.0), 0.0)
            return (left, top, right, bottom)

        w = _safe_float(rect.get('width', 0.0), 0.0)
        h = _safe_float(rect.get('height', 0.0), 0.0)
        return (left, top, left + w, top + h)

    return None


def _median(values: Sequence[float]) -> float:
    vals = [float(v) for v in values if v is not None]
    if not vals:
        return 0.0
    vals.sort()
    n = len(vals)
    mid = n // 2
    if (n % 2) == 1:
        return float(vals[mid])
    return (float(vals[mid - 1]) + float(vals[mid])) / 2.0


def _join_text(a: str, b: str) -> str:
    aa = (a or '').rstrip()
    bb = (b or '').lstrip()
    if not aa:
        return bb
    if not bb:
        return aa

    # Якщо попереднє закінчується дефісом, часто треба зшивати без пробілу.
    if aa.endswith('-'):
        return aa + bb

    # Якщо наступне починається пунктуацією, пробіл не потрібен.
    if bb[:1] in [',', '.', ';', ':', '!', '?', ')', ']', '}', '…']:
        return aa + bb

    # Якщо попереднє закінчується відкриваючою дужкою, пробіл теж не треба.
    if aa.endswith('(') or aa.endswith('[') or aa.endswith('{'):
        return aa + bb

    return aa + ' ' + bb


@dataclass
class _Box:
    text: str
    l: float
    t: float
    r: float
    b: float
    confidence: float
    bg_color: Optional[List[int]]

    @property
    def w(self) -> float:
        return max(0.0, float(self.r - self.l))

    @property
    def h(self) -> float:
        return max(0.0, float(self.b - self.t))

    @property
    def cx(self) -> float:
        return (float(self.l) + float(self.r)) / 2.0

    @property
    def cy(self) -> float:
        return (float(self.t) + float(self.b)) / 2.0


def merge_close_text_regions(
    regions: List[TextRegion],
    enabled: bool = True,
    x_gap_ratio: float = 1.25,
    line_y_ratio: float = 0.70,
    merge_vertical: bool = True,
) -> List[TextRegion]:
    """Зшиває близькі OCR-регіони в один рядок (word/segment merge).

    Ідея проста: OCR часто повертає кілька прямокутників на один рядок (або навіть слово).
    Переклад по шматках дає кашу. Тут ми групуємо по рядках (Y) і зшиваємо по X.

    За замовчуванням НЕ зливаємо рядки вертикально, тільки сегменти в межах рядка.
    """

    if not enabled:
        return regions

    src = regions or []
    if len(src) <= 1:
        return src

    boxes: List[_Box] = []
    heights: List[float] = []

    for r in src:
        rect = _rect_to_ltrb(getattr(r, 'rect', None))
        if rect is None:
            continue
        l, t, rr, b = rect
        if rr <= l or b <= t:
            continue
        text = (getattr(r, 'text', '') or '').strip()
        if not text:
            continue
        conf = _safe_float(getattr(r, 'confidence', 0.0), 0.0)
        bg = getattr(r, 'bg_color', None)
        bg2: Optional[List[int]] = None
        if isinstance(bg, (list, tuple)) and len(bg) == 3:
            try:
                bg2 = [int(bg[0]), int(bg[1]), int(bg[2])]
            except Exception:
                bg2 = None
        bx = _Box(text=text, l=float(l), t=float(t), r=float(rr), b=float(b), confidence=conf, bg_color=bg2)
        boxes.append(bx)
        heights.append(bx.h)

    if len(boxes) <= 1:
        return regions

    med_h = _median(heights)
    if med_h <= 0.0:
        med_h = float(max(10.0, max(heights or [10.0])))

    # Пороги
    max_gap_x = max(2.0, float(med_h) * float(x_gap_ratio))
    max_dy_line = max(2.0, float(med_h) * float(line_y_ratio))

    # 1) Групуємо по рядках (Y)
    boxes.sort(key=lambda b: (b.cy, b.cx))

    line_groups: List[List[_Box]] = []
    line_centers: List[float] = []

    for bx in boxes:
        best_i = -1
        best_dy = 10 ** 18

        for i, cy in enumerate(line_centers):
            dy = abs(float(bx.cy) - float(cy))
            if dy <= max_dy_line and dy < best_dy:
                best_dy = dy
                best_i = i

        if best_i < 0:
            line_groups.append([bx])
            line_centers.append(float(bx.cy))
        else:
            line_groups[best_i].append(bx)
            # оновлюємо центр рядка
            n = float(len(line_groups[best_i]))
            line_centers[best_i] = (float(line_centers[best_i]) * (n - 1.0) + float(bx.cy)) / n

    # 2) В кожному рядку зливаємо по X
    merged: List[_Box] = []

    for group in line_groups:
        group.sort(key=lambda b: (b.l, b.t))
        cur: Optional[_Box] = None

        for bx in group:
            if cur is None:
                cur = _Box(
                    text=bx.text,
                    l=bx.l,
                    t=bx.t,
                    r=bx.r,
                    b=bx.b,
                    confidence=bx.confidence,
                    bg_color=bx.bg_color,
                )
                continue

            gap = float(bx.l - cur.r)
            overlap_y = float(min(cur.b, bx.b) - max(cur.t, bx.t))
            min_h = max(1.0, min(cur.h, bx.h))
            overlap_ratio = float(overlap_y) / float(min_h)

            # схожість висот (щоб не зшивати різні «ряди» або UI-елементи)
            h_ratio = float(max(cur.h, bx.h)) / float(max(1.0, min(cur.h, bx.h)))

            can_merge = True
            if gap > max_gap_x:
                can_merge = False
            if overlap_ratio < 0.35:
                can_merge = False
            if h_ratio > 1.9:
                can_merge = False

            if can_merge:
                cur.text = _join_text(cur.text, bx.text)
                cur.l = float(min(cur.l, bx.l))
                cur.t = float(min(cur.t, bx.t))
                cur.r = float(max(cur.r, bx.r))
                cur.b = float(max(cur.b, bx.b))

                # confidence: грубо середнє
                cur.confidence = float((cur.confidence + bx.confidence) / 2.0)

                # bg_color: якщо обидва є, середнє
                if cur.bg_color is not None and bx.bg_color is not None:
                    try:
                        cur.bg_color = [
                            int((int(cur.bg_color[0]) + int(bx.bg_color[0])) / 2),
                            int((int(cur.bg_color[1]) + int(bx.bg_color[1])) / 2),
                            int((int(cur.bg_color[2]) + int(bx.bg_color[2])) / 2),
                        ]
                    except Exception:
                        pass
                elif cur.bg_color is None and bx.bg_color is not None:
                    cur.bg_color = bx.bg_color

            else:
                merged.append(cur)
                cur = _Box(
                    text=bx.text,
                    l=bx.l,
                    t=bx.t,
                    r=bx.r,
                    b=bx.b,
                    confidence=bx.confidence,
                    bg_color=bx.bg_color,
                )

        if cur is not None:
            merged.append(cur)

    # 3) (опційно) зливаємо сусідні рядки в один мультилайн-бокс
    merged2: List[_Box] = merged
    if bool(merge_vertical) and len(merged2) > 1:
        merged2.sort(key=lambda b: (b.t, b.l))
        out_v: List[_Box] = []
        cur_v: Optional[_Box] = None

        max_v_gap = max(2.0, float(med_h) * 0.90)
        min_x_overlap_ratio = 0.55
        align_tol = float(med_h) * 1.25

        for bx in merged2:
            if cur_v is None:
                cur_v = _Box(
                    text=bx.text,
                    l=bx.l,
                    t=bx.t,
                    r=bx.r,
                    b=bx.b,
                    confidence=bx.confidence,
                    bg_color=bx.bg_color,
                )
                continue

            v_gap = float(bx.t - cur_v.b)
            if v_gap < -max_v_gap:
                # перекриття по Y занадто велике/дивне, не зшиваємо
                v_gap = max_v_gap + 1.0

            x_overlap = float(min(cur_v.r, bx.r) - max(cur_v.l, bx.l))
            min_w = max(1.0, min(cur_v.w, bx.w))
            x_overlap_ratio = float(x_overlap) / float(min_w)

            left_aligned = abs(float(cur_v.l) - float(bx.l)) <= align_tol
            center_aligned = abs(float(cur_v.cx) - float(bx.cx)) <= align_tol

            can_merge_v = True
            if v_gap > max_v_gap:
                can_merge_v = False
            if x_overlap_ratio < min_x_overlap_ratio:
                can_merge_v = False
            if not (left_aligned or center_aligned):
                can_merge_v = False

            if can_merge_v:
                cur_v.text = (cur_v.text or '').rstrip() + "\n" + (bx.text or '').lstrip()
                cur_v.l = float(min(cur_v.l, bx.l))
                cur_v.t = float(min(cur_v.t, bx.t))
                cur_v.r = float(max(cur_v.r, bx.r))
                cur_v.b = float(max(cur_v.b, bx.b))
                cur_v.confidence = float((cur_v.confidence + bx.confidence) / 2.0)

                if cur_v.bg_color is not None and bx.bg_color is not None:
                    try:
                        cur_v.bg_color = [
                            int((int(cur_v.bg_color[0]) + int(bx.bg_color[0])) / 2),
                            int((int(cur_v.bg_color[1]) + int(bx.bg_color[1])) / 2),
                            int((int(cur_v.bg_color[2]) + int(bx.bg_color[2])) / 2),
                        ]
                    except Exception:
                        pass
                elif cur_v.bg_color is None and bx.bg_color is not None:
                    cur_v.bg_color = bx.bg_color
            else:
                out_v.append(cur_v)
                cur_v = _Box(
                    text=bx.text,
                    l=bx.l,
                    t=bx.t,
                    r=bx.r,
                    b=bx.b,
                    confidence=bx.confidence,
                    bg_color=bx.bg_color,
                )

        if cur_v is not None:
            out_v.append(cur_v)

        merged2 = out_v

    # 4) Повертаємо як TextRegion
    out: List[TextRegion] = []
    for bx in merged2:
        rect: Dict[str, int] = {
            'left': int(round(bx.l)),
            'top': int(round(bx.t)),
            'right': int(round(bx.r)),
            'bottom': int(round(bx.b)),
        }
        out.append(
            TextRegion(
                text=str(bx.text),
                rect=rect,
                confidence=float(bx.confidence),
                is_dialog=False,
                bg_color=bx.bg_color,
            )
        )

    # Стабільний порядок: згори-вниз, зліва-направо
    out.sort(key=lambda r: (
        _safe_float((r.rect or {}).get('top', 0), 0.0),
        _safe_float((r.rect or {}).get('left', 0), 0.0),
    ))
    return out
