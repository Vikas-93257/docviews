"""Image facts, read from the bytes — PNG, JPEG, GIF, BMP, TIFF.

``add_picture`` has to know two things about an image: how big it is, and what
resolution it claims, because Word sizes an image in *inches* determined by
pixels-per-inch. Both live in the file header, so they are read here rather
than by pulling in Pillow: this library stays dependency-free (census §6 is
built on that) and a picture that cannot be measured is refused with a message
instead of being inserted at a guessed size.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple

EMU_PER_INCH = 914400
DEFAULT_DPI = 72.0            # what python-docx assumes without a resolution


class ImageError(ValueError):
    """The bytes are not an image this library can measure."""


@dataclass(frozen=True)
class ImageInfo:
    content_type: str          # image/png, image/jpeg, …
    extension: str             # png, jpeg, gif, bmp, tiff
    width_px: int
    height_px: int
    dpi_x: float = DEFAULT_DPI
    dpi_y: float = DEFAULT_DPI

    @property
    def native_width(self) -> int:
        """Width in EMU at the image's own resolution."""
        return int(round(self.width_px / self.dpi_x * EMU_PER_INCH))

    @property
    def native_height(self) -> int:
        return int(round(self.height_px / self.dpi_y * EMU_PER_INCH))

    def scaled(self, width: Optional[int] = None,
               height: Optional[int] = None) -> Tuple[int, int]:
        """(width, height) in EMU, keeping the image's aspect ratio.

        Word's behaviour, and python-docx's: give one side, the other follows;
        give neither, the native size is used.
        """
        native_w, native_h = self.native_width, self.native_height
        if width is None and height is None:
            return native_w, native_h
        ratio = self.width_px / self.height_px if self.height_px else 1.0
        if width is not None and height is not None:
            return int(width), int(height)
        if width is not None:
            return int(width), int(round(width / ratio))
        return int(round(height * ratio)), int(height)


# ------------------------------------------------------------------- detection
_SIGNATURE = 8                # enough bytes to tell every format apart


def read_source(source) -> Tuple[bytes, str]:
    """Bytes and a filename, from a path, bytes, or a file object.

    ``add_picture`` accepts all three, the way python-docx does.
    """
    if isinstance(source, (bytes, bytearray, memoryview)):
        return bytes(source), "image"
    if isinstance(source, (str, Path)):
        path = Path(source)
        return path.read_bytes(), path.name
    name = getattr(source, "name", "") or "image"
    if hasattr(source, "read"):
        data = source.read()
        if isinstance(data, str):
            data = data.encode()
        return bytes(data), str(name)
    raise ImageError(f"cannot read an image from {type(source).__name__}")


def detect(source) -> ImageInfo:
    """Measure an image: a path, bytes, or a file object."""
    data, _name = read_source(source)
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return _png(data)
    if data[:2] == b"\xff\xd8":
        return _jpeg(data)
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return _gif(data)
    if data[:2] == b"BM":
        return _bmp(data)
    if data[:4] in (b"II*\x00", b"MM\x00*"):
        return _tiff(data)
    if data[:5] == b"<?xml" or data[:4] == b"<svg":
        raise ImageError(
            "SVG cannot be embedded in a .docx directly: Word needs a raster "
            "fallback. Convert it to PNG first (or pass a PNG/JPEG/GIF/BMP/TIFF).")
    raise ImageError(
        "unrecognised image format: this library measures PNG, JPEG, GIF, BMP "
        "and TIFF without dependencies, and will not guess a size for anything "
        "else")


def _png(data: bytes) -> ImageInfo:
    if len(data) < 33:
        raise ImageError("PNG truncated before its header")
    width, height = struct.unpack(">II", data[16:24])
    dpi_x = dpi_y = DEFAULT_DPI
    offset = 8
    while offset + 8 <= len(data):
        length, kind = struct.unpack(">I4s", data[offset:offset + 8])
        if kind == b"pHYs" and offset + 8 + 9 <= len(data):
            x_ppu, y_ppu, unit = struct.unpack(">IIB", data[offset + 8:offset + 17])
            if unit == 1 and x_ppu and y_ppu:            # pixels per metre
                # python-docx exposes integer DPI, rounding pixels-per-metre
                # conversion before calculating the native EMU size.
                dpi_x = round(x_ppu * 0.0254)
                dpi_y = round(y_ppu * 0.0254)
            break
        if kind == b"IDAT" or kind == b"IEND":
            break
        offset += 12 + length
    return ImageInfo("image/png", "png", width, height,
                     _sane_dpi(dpi_x), _sane_dpi(dpi_y))


def _jpeg(data: bytes) -> ImageInfo:
    offset, dpi_x, dpi_y = 2, DEFAULT_DPI, DEFAULT_DPI
    width = height = 0
    while offset + 4 <= len(data):
        if data[offset] != 0xFF:
            offset += 1
            continue
        marker = data[offset + 1]
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            offset += 2
            continue
        if offset + 4 > len(data):
            break
        length = struct.unpack(">H", data[offset + 2:offset + 4])[0]
        segment = data[offset + 4:offset + 2 + length]
        if marker == 0xE0 and segment[:5] == b"JFIF\x00" and len(segment) >= 12:
            units = segment[7]
            if units:                                    # 1 = dpi, 2 = dots/cm
                x, y = struct.unpack(">HH", segment[8:12])
                if x and y:
                    dpi_x, dpi_y = (x, y) if units == 1 else (x * 2.54, y * 2.54)
        elif 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
            if len(segment) >= 5:
                height, width = struct.unpack(">HH", segment[1:5])
            break
        elif marker == 0xDA:                             # start of scan: no header left
            break
        offset += 2 + length
    if not width or not height:
        raise ImageError("JPEG has no frame header (SOF marker)")
    return ImageInfo("image/jpeg", "jpeg", width, height,
                     _sane_dpi(dpi_x), _sane_dpi(dpi_y))


def _gif(data: bytes) -> ImageInfo:
    if len(data) < 10:
        raise ImageError("GIF truncated before its screen descriptor")
    width, height = struct.unpack("<HH", data[6:10])
    return ImageInfo("image/gif", "gif", width, height)


def _bmp(data: bytes) -> ImageInfo:
    if len(data) < 26:
        raise ImageError("BMP truncated before its info header")
    header_size = struct.unpack("<I", data[14:18])[0]
    if header_size == 12:                                # BITMAPCOREHEADER
        width, height = struct.unpack("<HH", data[18:22])
    else:
        width, height = struct.unpack("<ii", data[18:26])
        height = abs(height)
    if not width or not height:
        raise ImageError("BMP reports a zero dimension")
    if len(data) < 46:
        raise ImageError("BMP truncated before its resolution fields")
    horz_px_per_meter, vert_px_per_meter = struct.unpack("<ii", data[38:46])
    dpi_x = int(round(horz_px_per_meter * 0.0254)) if horz_px_per_meter else 96
    dpi_y = int(round(vert_px_per_meter * 0.0254)) if vert_px_per_meter else 96
    return ImageInfo("image/bmp", "bmp", width, height, dpi_x, dpi_y)


def _tiff(data: bytes) -> ImageInfo:
    endian = "<" if data[:2] == b"II" else ">"
    magic = struct.unpack(endian + "H", data[2:4])[0]
    if magic != 42:
        raise ImageError("TIFF header is not a classic TIFF")
    offset = struct.unpack(endian + "I", data[4:8])[0]
    if offset + 2 > len(data):
        raise ImageError("TIFF points past the end of the file")
    count = struct.unpack(endian + "H", data[offset:offset + 2])[0]
    # tag -> value for the single-valued types this reader needs: SHORT, LONG,
    # RATIONAL. A value that does not fit in the four inline bytes lives at an
    # offset — which is where a resolution always lives.
    widths = {3: ("H", 2), 4: ("I", 4), 5: ("II", 8), 12: ("d", 8)}
    values = {}
    for index in range(count):
        entry = offset + 2 + index * 12
        if entry + 12 > len(data):
            break
        tag, kind, number = struct.unpack(endian + "HHI", data[entry:entry + 8])
        if kind not in widths or number != 1:
            continue
        fmt, size = widths[kind]
        raw = data[entry + 8:entry + 12]
        if size > 4:
            start = struct.unpack(endian + "I", raw)[0]
        else:
            start = entry + 8
        try:
            parts = struct.unpack(endian + fmt, data[start:start + size])
        except struct.error:
            continue
        if kind == 5:
            top, bottom = parts
            values[tag] = top / bottom if bottom else 0.0
        else:
            values[tag] = parts[0]
    width, height = values.get(256), values.get(257)
    if not width or not height:
        raise ImageError("TIFF has no image dimensions")
    dpi_x = dpi_y = DEFAULT_DPI
    unit = values.get(296, 2)                            # 2 = inches, 3 = cm
    if values.get(282) and values.get(283):
        dpi_x, dpi_y = values[282], values[283]
        if unit == 3:                                    # centimetres
            dpi_x, dpi_y = dpi_x * 2.54, dpi_y * 2.54
    return ImageInfo("image/tiff", "tiff", width, height,
                     _sane_dpi(dpi_x), _sane_dpi(dpi_y))


def _sane_dpi(value: float) -> float:
    """A resolution Word will believe: 1..2400, else the 72 dpi default."""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return DEFAULT_DPI
    return number if 1.0 <= number <= 2400.0 else DEFAULT_DPI
