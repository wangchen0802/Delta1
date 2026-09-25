"""Minimal Embedded OpenType (EOT v2.1, uncompressed) writer.

PowerPoint stores embedded fonts as ppt/fonts/*.fntdata, which are EOT files
wrapping a plain TrueType font. This mirrors the header layout the original
SimReal deck used (Version 0x00020001, Flags 0, no MTX compression / XOR).
"""
import struct
from fontTools.ttLib import TTFont


def _name(font, name_id):
    rec = font["name"].getName(name_id, 3, 1, 0x409) or font["name"].getName(name_id, 1, 0, 0)
    return rec.toUnicode() if rec else ""


def _string_field(s):
    b = s.encode("utf-16-le")
    return struct.pack("<H", len(b)) + b


def ttf_to_eot(ttf_bytes, font=None):
    import io
    font = font or TTFont(io.BytesIO(ttf_bytes))
    os2 = font["OS/2"]
    head = font["head"]
    panose = os2.panose
    panose_bytes = bytes([
        panose.bFamilyType, panose.bSerifStyle, panose.bWeight, panose.bProportion,
        panose.bContrast, panose.bStrokeVariation, panose.bArmStyle, panose.bLetterForm,
        panose.bMidline, panose.bXHeight,
    ])
    italic = 1 if (os2.fsSelection & 1) else 0
    ur = os2.ulUnicodeRange1, os2.ulUnicodeRange2, os2.ulUnicodeRange3, os2.ulUnicodeRange4
    cp = getattr(os2, "ulCodePageRange1", 0), getattr(os2, "ulCodePageRange2", 0)

    body = b"".join([
        struct.pack("<I", 0x00020001),          # Version
        struct.pack("<I", 0),                   # Flags
        panose_bytes,
        struct.pack("<BB", 1, italic),          # Charset (DEFAULT_CHARSET), Italic
        struct.pack("<I", os2.usWeightClass),
        struct.pack("<H", os2.fsType),
        struct.pack("<H", 0x504C),              # MagicNumber
        struct.pack("<IIII", *ur),
        struct.pack("<II", *cp),
        struct.pack("<I", head.checkSumAdjustment),
        struct.pack("<IIII", 0, 0, 0, 0),       # Reserved1-4
        struct.pack("<H", 0), _string_field(_name(font, 1)),   # Padding1, FamilyName
        struct.pack("<H", 0), _string_field(_name(font, 2)),   # Padding2, StyleName
        struct.pack("<H", 0), _string_field(_name(font, 5)),   # Padding3, VersionName
        struct.pack("<H", 0), _string_field(_name(font, 4)),   # Padding4, FullName
        struct.pack("<H", 0), struct.pack("<H", 0),            # Padding5, RootStringSize=0
    ])
    font_data_size = len(ttf_bytes)
    eot_size = 8 + len(body) + font_data_size
    return struct.pack("<II", eot_size, font_data_size) + body + ttf_bytes


def eot_to_ttf(eot_bytes):
    eot_size, font_data_size = struct.unpack("<II", eot_bytes[:8])
    return eot_bytes[-font_data_size:]
