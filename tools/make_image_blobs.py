"""Regenerate ``tests/blobs.py`` — the tiny images the image tests use.

Needs Pillow, and only for maintenance: the suite itself imports nothing but
base64, so a contributor without Pillow can still run every test.

    python3 tools/make_image_blobs.py [tests/blobs.py]
"""

from __future__ import annotations

import base64
import io
import sys

from PIL import Image


def blob(size, fmt, dpi=None, **kwargs) -> bytes:
    image = Image.new("RGB", size, (200, 30, 30))
    buffer = io.BytesIO()
    if dpi:
        kwargs["dpi"] = dpi
    image.save(buffer, format=fmt, **kwargs)
    return buffer.getvalue()


BLOBS = {
    "PNG_96": blob((120, 60), "PNG", dpi=(96, 96)),
    "PNG_RAW": blob((120, 60), "PNG"),
    "JPG_300": blob((200, 100), "JPEG", dpi=(300, 300), quality=60),
    "GIF": blob((3, 2), "GIF"),
    "BMP": blob((4, 6), "BMP"),
    "TIFF_300": blob((5, 7), "TIFF", dpi=(300, 300)),
    "SVG": b'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>',
    "JUNK": b"not an image at all",
}


def render(blobs) -> str:
    out = ['"""Tiny image fixtures, base64-embedded so the suite needs no image library.',
           '',
           'Regenerate with::',
           '',
           '    python3 tools/make_image_blobs.py',
           '"""',
           '',
           'import base64',
           '',
           '']
    for name, data in blobs.items():
        encoded = base64.b64encode(data).decode()
        out.append(f"{name} = base64.b64decode(")
        out.extend(f'    "{encoded[i:i + 68]}"' for i in range(0, len(encoded), 68))
        out.append(")")
        out.append("")
    out.extend([
        "class _Named:",
        "    def __init__(self, name, data):",
        "        self.name, self._data = name, data",
        "",
        "    def read(self):",
        "        return self._data",
        "",
        "    def seek(self, offset, whence=0):",
        "        return 0",
        "",
        "",
        "def file_like(name: str, data: bytes):",
        '    """A file object with a ``.name``, like the ones callers pass in."""',
        "    return _Named(name, data)",
        "",
        "",
        "BLOBS = {",
    ])
    out.extend(f'    "{name}": {name},' for name in blobs)
    out.append("}")
    return "\n".join(out) + "\n"


def main() -> int:
    path = sys.argv[1] if len(sys.argv) > 1 else "tests/blobs.py"
    with open(path, "w") as handle:
        handle.write(render(BLOBS))
    print(f"wrote {path} ({len(BLOBS)} fixtures)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
