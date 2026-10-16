"""
Positional replacement of the cached values inside a chart series block.

Swapping values by string search ("replace the first <c:v>8</c:v>") puts numbers
in the wrong place whenever old and new values overlap - e.g. old [130, 130, 104]
-> new [130, 104, 86] rewrites the first 130 twice. These helpers rewrite the
numCache by point index instead, so position i always gets value i.
"""
import re


def set_cache(block, tag, values):
    """Rewrite the <c:numCache>/<c:numLit> inside the first <tag>...</tag> of `block`
    (tag is 'c:val', 'c:xVal' or 'c:yVal') to hold `values`, in order.
    formatCode and any trailing extLst are kept; ptCount follows len(values)."""
    m = re.search(rf"<{tag}>.*?</{tag}>", block, re.DOTALL)
    if not m:
        raise ValueError(f"no <{tag}> in series block")
    part = m.group(0)
    c = re.search(r"(<c:(numCache|numLit)>)(.*?)(</c:\2>)", part, re.DOTALL)
    if not c:
        raise ValueError(f"no numCache/numLit in <{tag}>")
    inner = c.group(3)
    fmt = re.search(r"<c:formatCode>.*?</c:formatCode>", inner, re.DOTALL)
    ext = re.search(r"<c:extLst>.*?</c:extLst>", inner, re.DOTALL)
    pts = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(values))
    new_inner = ((fmt.group(0) if fmt else "") + f'<c:ptCount val="{len(values)}"/>'
                 + pts + (ext.group(0) if ext else ""))
    new_part = part[:c.start(3)] + new_inner + part[c.end(3):]
    return block.replace(part, new_part, 1)


def cache_values(block, tag):
    """Values of the first <tag> numCache/numLit in `block`, by point index."""
    m = re.search(rf"<{tag}>.*?</{tag}>", block, re.DOTALL)
    if not m:
        return []
    pts = re.findall(r'<c:pt idx="(\d+)"[^>]*>\s*<c:v>([^<]*)</c:v>', m.group(0))
    out = [None] * (max((int(i) for i, _ in pts), default=-1) + 1)
    for i, v in pts:
        out[int(i)] = v
    return out


def norm(name):
    return re.sub(r"[^a-z0-9]", "", name.lower())
