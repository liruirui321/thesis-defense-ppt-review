# PPTX Python Checks

Use `python-pptx` for structural checks. Always wrap shape coordinates in safe accessors because some PPTX files contain float coordinate strings that can break direct `int()` conversion.

```python
from pptx import Presentation

prs = Presentation("deck.pptx")

def safe(sh, attr):
    try:
        return int(getattr(sh, attr))
    except Exception:
        return None

def tx(sh):
    try:
        if sh.has_text_frame:
            return " ".join(
                "".join(run.text for run in p.runs)
                for p in sh.text_frame.paragraphs
            ).strip()
    except Exception:
        return ""
    return ""
```

## Duplicate References

```python
for si, slide in enumerate(prs.slides, 1):
    refs = []
    for j, sh in enumerate(slide.shapes):
        text = tx(sh)
        if text.startswith("参考："):
            refs.append((j, text))
    if len(refs) > 1:
        print(si, refs)
```

## Body Centering

Exclude title, source lines, page numbers, logos, style rules, and footers. Compute body bounding boxes and flag large offsets.

```python
W = int(prs.slide_width) / 914400
DES_Y = 3.70

for si, slide in enumerate(prs.slides, 1):
    xs, ys = [], []
    for sh in slide.shapes:
        l, t, w, h = [safe(sh, a) for a in ("left", "top", "width", "height")]
        if None in (l, t, w, h):
            continue
        x, y, ww, hh = l/914400, t/914400, w/914400, h/914400
        text = tx(sh)
        name = getattr(sh, "name", "")
        if name.startswith(("REF_STYLE", "REF_SECTION")):
            continue
        if y < 1.05 or y > 6.35:
            continue
        if text.startswith("参考："):
            continue
        if text and text.isdigit() and ww < .8:
            continue
        if ww * hh < .08 or hh < .05:
            continue
        xs += [x, x + ww]
        ys += [y, y + hh]
    if xs and ys:
        xoff = (min(xs) + max(xs)) / 2 - W / 2
        ycenter = (min(ys) + max(ys)) / 2
        if abs(xoff) > .18 or abs(ycenter - DES_Y) > .45:
            print(si, "xoff", round(xoff, 2), "ycenter", round(ycenter, 2))
```

## Overlap And Footer Checks

```python
def rect(sh):
    l, t, w, h = [safe(sh, a) for a in ("left", "top", "width", "height")]
    if None in (l, t, w, h):
        return None
    return l, t, l+w, t+h

def area(r):
    return max(0, r[2]-r[0]) * max(0, r[3]-r[1])

def inter(a, b):
    x1, y1 = max(a[0], b[0]), max(a[1], b[1])
    x2, y2 = min(a[2], b[2]), min(a[3], b[3])
    return max(0, x2-x1) * max(0, y2-y1)

issues = []
for si, slide in enumerate(prs.slides, 1):
    nums, logos, texts, pics = [], [], [], []
    for sh in slide.shapes:
        r = rect(sh)
        if not r:
            continue
        l, t, rr, b = r
        text = tx(sh)
        if si != 1 and text in {str(si), f"{si:02d}"} and l > int(11.8*914400) and t > int(6.8*914400):
            nums.append(r)
        if sh.shape_type == 13 and t < int(.8*914400) and l > int(9*914400):
            logos.append(r)
        if text and int(.95*914400) < t < int(6.25*914400) and len(text) > 1:
            texts.append((r, text[:35]))
        if sh.shape_type == 13 and int(.95*914400) < t < int(6.25*914400):
            pics.append((r, sh.name))
    if si != 1 and len(nums) != 1:
        issues.append((si, "bottom_page_number_count", len(nums)))
    if si != 1 and len(logos) < 1:
        issues.append((si, "logo_missing"))
    for i in range(len(texts)):
        for j in range(i+1, len(texts)):
            ov = inter(texts[i][0], texts[j][0])
            if ov and ov / min(area(texts[i][0]), area(texts[j][0])) > .42:
                issues.append((si, "text_overlap", texts[i][1], texts[j][1]))
print(issues)
```

## Top Band And Section Chip Checks

Flag non-white title bands and redundant top-right section chips such as "第七部分  总结展望".

```python
non_white_top_bands = []
section_chips = []

for si, slide in enumerate(prs.slides, 1):
    for j, sh in enumerate(slide.shapes):
        text = tx(sh)
        l, t, w, h = [safe(sh, a) for a in ("left", "top", "width", "height")]
        if None in (l, t, w, h):
            continue
        if "第" in text and "部分" in text and t < int(.7 * 914400):
            section_chips.append((si, j, text))
        try:
            if (
                t <= int(.05 * 914400)
                and l <= int(.05 * 914400)
                and w >= int(12.5 * 914400)
                and int(.9 * 914400) <= h <= int(1.4 * 914400)
            ):
                rgb = str(sh.fill.fore_color.rgb) if sh.fill.type == 1 else None
                if rgb != "FFFFFF":
                    non_white_top_bands.append((si, j, rgb))
        except Exception:
            pass

print("non_white_top_bands", non_white_top_bands)
print("section_chips", section_chips)
```

## Connector Intrusion Checks

Check lines against both text boxes and visible rectangles. For connector-like lines, zero-height shapes need a small hitbox because the rendered stroke still has thickness.

```python
def bbox(sh, pad=0):
    l, t, w, h = [safe(sh, a) for a in ("left", "top", "width", "height")]
    if None in (l, t, w, h):
        return None
    return l - pad, t - pad, l + w + pad, t + h + pad

connector_issues = []

for si, slide in enumerate(prs.slides, 1):
    lines, rects, texts = [], [], []
    for j, sh in enumerate(slide.shapes):
        l, t, w, h = [safe(sh, a) for a in ("left", "top", "width", "height")]
        if None in (l, t, w, h):
            continue
        y = t / 914400
        xw = w / 914400
        yh = h / 914400
        text = tx(sh)
        if sh.shape_type == 9 and 1.2 < y < 6.2:
            lines.append((j, bbox(sh, int(.015 * 914400))))
        if sh.shape_type == 1 and 1.2 < y < 6.2 and xw > .3 and yh > .08:
            rects.append((j, bbox(sh), text[:35]))
        if text and 1.2 < y < 6.2:
            texts.append((j, bbox(sh), text[:35]))
    for li, lr in lines:
        if not lr:
            continue
        for ri, rr, rt in rects:
            if rr and inter(lr, rr):
                connector_issues.append((si, "line_rect", li, ri, rt))
        for ti, tr, tt in texts:
            if tr and inter(lr, tr):
                connector_issues.append((si, "line_text", li, ti, tt))

print("connector_issues", connector_issues)
```

## Card Text Containment Spot Check

Use this for slides with explanatory cards or callout panels. It checks independent text boxes against blank background rectangles. Text embedded directly in an auto-shape is treated as self-contained to avoid false positives in tables and diagrams. Verify findings visually.

```python
def contains(outer, inner, margin=0):
    return (
        outer[0] - margin <= inner[0]
        and outer[1] - margin <= inner[1]
        and outer[2] + margin >= inner[2]
        and outer[3] + margin >= inner[3]
    )

containment_issues = []

for si, slide in enumerate(prs.slides, 1):
    rects, texts = [], []
    for j, sh in enumerate(slide.shapes):
        l, t, w, h = [safe(sh, a) for a in ("left", "top", "width", "height")]
        if None in (l, t, w, h):
            continue
        x, y, ww, hh = l / 914400, t / 914400, w / 914400, h / 914400
        text = tx(sh)
        if sh.shape_type == 1 and 1.1 < y < 6.6 and ww > 1 and hh > .5 and not text:
            rects.append((j, rect(sh)))
        if (
            sh.shape_type == 17
            and text
            and 1.1 < y < 6.6
            and not text.startswith("参考：")
            and not text.isdigit()
        ):
            texts.append((j, rect(sh), text[:45]))
    for ti, tr, tt in texts:
        for ri, rr in rects:
            if (
                rr[0] - int(.03 * 914400) <= tr[0] <= rr[2] + int(.03 * 914400)
                and rr[1] - int(.03 * 914400) <= tr[1] <= rr[3] + int(.03 * 914400)
            ):
                if not contains(rr, tr, int(.03 * 914400)):
                    containment_issues.append((si, ti, ri, tt))
                break

print("containment_issues", containment_issues)
```

## Long Shallow Text Risk Check

Use this after fixing any text-overflow problem. It catches long text in shallow text boxes, auto-shapes, conclusion bars, and metric cards. These are often visible clipping risks even when no separate background rectangle exists.

```python
long_shallow = []

for si, slide in enumerate(prs.slides, 1):
    for j, sh in enumerate(slide.shapes):
        text = tx(sh).replace("\n", " ")
        l, t, w, h = [safe(sh, a) for a in ("left", "top", "width", "height")]
        if None in (l, t, w, h):
            continue
        y = t / 914400
        hh = h / 914400
        if len(text) >= 34 and 1.1 < y < 6.5 and hh < .34:
            long_shallow.append((si, j, round(l / 914400, 2), round(y, 2), round(w / 914400, 2), round(hh, 2), text[:80]))

print("long_shallow_text", long_shallow)
```
