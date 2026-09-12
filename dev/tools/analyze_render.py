"""Measure a rendered preview PNG: is it paginated, and where is the paper edge?

Eyeballing screenshots is how this whole question got murky. This turns a render into
numbers, so "the preview has no page layout" becomes a checkable claim.

What it reports:
  content box      first/last column and first/last row containing ink
  gutter rows      runs of rows whose ink spans nearly the full width - that is what the
                   grey band between two Word pages looks like when a preview paginates
  ink density      share of inked pixels, which separates "one mostly-empty page" from
                   "continuous reflowed text"

Usage:
    python analyze_render.py <png> [--label "..."]
    python analyze_render.py <png> --rows      # dump the per-row ink profile
"""

import struct
import sys
import zlib


def read_png(path):
    """Decode an 8-bit truecolour PNG (what capture_preview.py and render_handler.py emit)."""
    with open(path, "rb") as f:
        data = f.read()

    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("not a PNG")

    pos = 8
    width = height = None
    idat = b""
    while pos < len(data):
        length = struct.unpack(">I", data[pos : pos + 4])[0]
        tag = data[pos + 4 : pos + 8]
        body = data[pos + 8 : pos + 8 + length]
        if tag == b"IHDR":
            width, height, depth, colour = struct.unpack(">IIBB", body[:10])
            if depth != 8 or colour != 2:
                raise ValueError(f"expected 8-bit truecolour, got depth={depth} colour={colour}")
        elif tag == b"IDAT":
            idat += body
        elif tag == b"IEND":
            break
        pos += 12 + length

    raw = zlib.decompress(idat)
    stride = width * 3
    rows = []
    prev = bytearray(stride)
    cursor = 0
    for _ in range(height):
        filter_type = raw[cursor]
        cursor += 1
        line = bytearray(raw[cursor : cursor + stride])
        cursor += stride
        # Undo the PNG scanline filter. The encoder only emits 0 (None), but be permissive.
        if filter_type == 1:
            for i in range(3, stride):
                line[i] = (line[i] + line[i - 3]) & 0xFF
        elif filter_type == 2:
            for i in range(stride):
                line[i] = (line[i] + prev[i]) & 0xFF
        elif filter_type == 3:
            for i in range(stride):
                left = line[i - 3] if i >= 3 else 0
                line[i] = (line[i] + ((left + prev[i]) >> 1)) & 0xFF
        elif filter_type == 4:
            for i in range(stride):
                a = line[i - 3] if i >= 3 else 0
                b = prev[i]
                c = prev[i - 3] if i >= 3 else 0
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                pred = a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)
                line[i] = (line[i] + pred) & 0xFF
        prev = line
        rows.append(bytes(line))
    return width, height, rows


def ink_profile(width, height, rows, threshold=245):
    """Per-row fraction of pixels that are not near-white."""
    profile = []
    for row in rows:
        inked = 0
        for x in range(0, width, 2):
            if (row[x * 3] < threshold or row[x * 3 + 1] < threshold
                    or row[x * 3 + 2] < threshold):
                inked += 1
        profile.append(inked / max(width // 2, 1))
    return profile


def analyse(path, dump_rows=False, label=None):
    width, height, rows = read_png(path)
    profile = ink_profile(width, height, rows)

    def col_ink(x):
        hits = 0
        offset = x * 3
        for y in range(0, height, 2):
            row = rows[y]
            if row[offset] < 245 or row[offset + 1] < 245 or row[offset + 2] < 245:
                hits += 1
        return hits

    col_has_ink = [col_ink(x) >= 1 for x in range(width)]
    left = next((x for x, v in enumerate(col_has_ink) if v), None)
    right = next((x for x in range(width - 1, -1, -1) if col_has_ink[x]), None)

    row_has_ink = [p > 0.002 for p in profile]
    top = next((y for y, v in enumerate(row_has_ink) if v), None)
    bottom = next((y for y in range(height - 1, -1, -1) if row_has_ink[y]), None)

    # A page gutter is a band of rows that are inked edge to edge. Text never does that;
    # the gap between two Word pages always does.
    gutter_rows = [y for y, p in enumerate(profile) if p > 0.90]
    runs = []
    for y in gutter_rows:
        if runs and y == runs[-1][1] + 1:
            runs[-1][1] = y
        else:
            runs.append([y, y])
    runs = [r for r in runs if r[1] - r[0] >= 4]

    print(f"{label or path}")
    print(f"   size            {width}x{height}")
    if left is not None:
        print(f"   content box     x {left}..{right}  y {top}..{bottom}")
        print(f"   side margins    left {left}px ({left / width:.1%})   "
              f"right {width - 1 - right}px ({(width - 1 - right) / width:.1%})")
    print(f"   mean ink        {sum(profile) / len(profile):.4f}")
    if runs:
        print(f"   full-width bands (= page gutters): {len(runs)}")
        for a, b in runs[:6]:
            print(f"       rows {a}..{b}  ({b - a + 1}px tall)")
    else:
        print("   full-width bands (= page gutters): NONE")

    if dump_rows:
        print("   row ink profile (every 20th row):")
        for y in range(0, height, 20):
            bar = "#" * int(profile[y] * 50)
            print(f"     {y:5d} {profile[y]:.3f} {bar}")

    return {"width": width, "height": height, "left": left, "right": right,
            "top": top, "bottom": bottom, "gutters": runs,
            "mean_ink": sum(profile) / len(profile)}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        return 1
    label = None
    if "--label" in sys.argv:
        label = sys.argv[sys.argv.index("--label") + 1]
    for path in args:
        analyse(path, dump_rows="--rows" in sys.argv, label=label)
        print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
