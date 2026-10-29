"""HWP(5.x)·HWPX 문서에서 본문 텍스트를 뽑는다. 표 안의 글자도 문서 순서대로 나온다.

사용 예:
    python hwp_text.py 제안요청서.hwp > 제안요청서.txt
"""

import struct
import sys
import zipfile
import zlib
import re

import olefile

PARA_TEXT = 67
# 컨트롤 문자 중 뒤따르는 7글자(14바이트)를 함께 차지하는 것
EXTENDED_CONTROLS = {1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23}


def hwp5_text(path):
    ole = olefile.OleFileIO(path)
    header = ole.openstream("FileHeader").read()
    compressed = bool(struct.unpack_from("<I", header, 36)[0] & 1)
    sections = sorted(
        (entry for entry in ole.listdir() if entry[0] == "BodyText"),
        key=lambda entry: int(entry[1].removeprefix("Section")),
    )
    lines = []
    for entry in sections:
        data = ole.openstream(entry).read()
        if compressed:
            data = zlib.decompress(data, -15)
        lines.extend(read_paragraphs(data))
    return "\n".join(lines)


def read_paragraphs(data):
    pos = 0
    while pos < len(data):
        (record,) = struct.unpack_from("<I", data, pos)
        tag, size = record & 0x3FF, (record >> 20) & 0xFFF
        pos += 4
        if size == 0xFFF:
            (size,) = struct.unpack_from("<I", data, pos)
            pos += 4
        if tag == PARA_TEXT:
            yield decode_para(data[pos:pos + size])
        pos += size


def decode_para(raw):
    chars = []
    i = 0
    while i + 1 < len(raw):
        (code,) = struct.unpack_from("<H", raw, i)
        if code == 9:
            chars.append("\t")
            i += 16
        elif code in EXTENDED_CONTROLS:
            i += 16
        elif code < 32:
            if code in (10, 13):
                chars.append(" ")
            i += 2
        else:
            chars.append(chr(code))
            i += 2
    return "".join(chars).rstrip()


def hwpx_text(path):
    lines = []
    with zipfile.ZipFile(path) as z:
        names = sorted(n for n in z.namelist() if re.match(r"Contents/section\d+\.xml", n))
        for name in names:
            xml = z.read(name).decode("utf-8")
            for para in re.findall(r"<hp:p\b.*?</hp:p>", xml, flags=re.S):
                text = "".join(re.findall(r"<hp:t>(.*?)</hp:t>", para, flags=re.S))
                lines.append(re.sub(r"<[^>]+>", "", text))
    return "\n".join(lines)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    path = sys.argv[1]
    text = hwpx_text(path) if path.lower().endswith(".hwpx") else hwp5_text(path)
    print(text)


if __name__ == "__main__":
    main()
