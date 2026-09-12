"""Build .docx fixtures that differ only in the stored view mode / zoom (word/settings.xml).

Word records the editing view and the zoom inside the document, and the Shell preview handler
is a Word renderer - so whether the preview paginates may depend on these two elements.
Making a fixture per value is how that gets settled without guessing.

<w:zoom> matters as much as <w:view> here:
  w:val="bestFit"   fit the page *width*   -> page fills the window, no visible paper edge
  w:val="fullPage"  fit the whole *page*   -> page shrinks, grey around it, page breaks show
  w:val="textFit"   fit the text width
  w:percent="NN"    explicit zoom percentage

IMPORTANT: <w:settings> is an XSD *sequence*, not a bag. w:view and w:zoom must appear in
their schema positions - near the very start, right after w:writeProtection, and view before
zoom. Appending them at the end of the element makes the part schema-invalid and Word then
silently ignores them, which looks exactly like "the setting has no effect". That mistake
already produced one round of misleading results, hence the explicit insertion point below.

Usage:
    python make_view_fixture.py <src.docx> <out.docx> <print|web|outline|normal|none>
                                [--zoom fullPage|bestFit|textFit|none|percent:NN]
"""

import os
import re
import sys
import zipfile

VIEWS = ("print", "web", "outline", "normal", "none")

VIEW_TAG = re.compile(r"<w:view\b[^>]*/>")
ZOOM_TAG = re.compile(r"<w:zoom\b[^>]*/>")
SETTINGS_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# w:view then w:zoom come first in CT_Settings after writeProtection.
ANCHORS = [
    re.compile(r"<w:writeProtection\b[^>]*/>"),
    re.compile(r"<w:writeProtection\b[^>]*>.*?</w:writeProtection>", re.S),
    re.compile(r"<w:settings\b[^>]*?(?:/>|>)"),
]

FALLBACK_SETTINGS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:settings xmlns:w="{SETTINGS_NS}"></w:settings>'
)


def zoom_tag(spec):
    if spec in (None, "keep", "none"):
        return None
    if spec in ("fullPage", "bestFit", "textFit"):
        return f'<w:zoom w:val="{spec}"/>'
    if spec.startswith("percent:"):
        return f'<w:zoom w:percent="{int(spec.split(":", 1)[1])}"/>'
    raise ValueError(f"unsupported zoom spec {spec!r}")


def _insert_early(xml, payload):
    """Insert payload at the schema-correct position: just after writeProtection (or right
    after the opening w:settings tag when there is none)."""
    for anchor in ANCHORS[:2]:
        match = anchor.search(xml)
        if match:
            return xml[: match.end()] + payload + xml[match.end():]

    # No writeProtection: insert right after the opening w:settings element.
    match = re.search(r"<w:settings\b[^>]*?/>", xml)
    if match:  # self-closing empty settings element
        return xml[: match.end() - 2] + ">" + payload + "</w:settings>" + xml[match.end():]

    match = re.search(r"<w:settings\b[^>]*>", xml)
    if not match:
        raise ValueError("no opening w:settings element")
    return xml[: match.end()] + payload + xml[match.end():]


def make(src, dst, view, zoom=None):
    if view not in VIEWS:
        raise ValueError(f"view must be one of {VIEWS}")

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        seen_settings = False
        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == "word/settings.xml":
                seen_settings = True
                xml = data.decode("utf-8")
                xml = VIEW_TAG.sub("", xml)
                xml = ZOOM_TAG.sub("", xml)

                payload = ""
                if view != "none":
                    payload += f'<w:view w:val="{view}"/>'
                tag = zoom_tag(zoom)
                if tag is not None:
                    payload += tag

                if payload:
                    xml = _insert_early(xml, payload)
                data = xml.encode("utf-8")

            zout.writestr(item, data)

        if not seen_settings:
            extra = ('<w:view w:val="%s"/>' % view) if view != "none" else ""
            extra += zoom_tag(zoom) or ""
            zout.writestr(
                "word/settings.xml",
                FALLBACK_SETTINGS.replace("</w:settings>", extra + "</w:settings>"),
            )

    # A fresh copy inherits nothing, but make sure no Zone.Identifier sneaks in.
    (lambda zone: os.remove(zone) if os.path.exists(zone) else None)(dst + ":Zone.Identifier")

    return os.path.abspath(dst)


def main():
    argv = sys.argv[1:]
    zoom = None
    positional = []
    i = 0
    while i < len(argv):
        if argv[i] == "--zoom":
            zoom = argv[i + 1]
            i += 2
        elif argv[i].startswith("--"):
            i += 1
        else:
            positional.append(argv[i])
            i += 1

    if len(positional) != 3:
        print(__doc__)
        return 1

    src, dst, view = positional
    out = make(src, dst, view, zoom)
    print(f"view={view:8s} zoom={str(zoom):12s} -> {out} ({os.path.getsize(out) / 1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

