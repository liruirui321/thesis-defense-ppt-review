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
