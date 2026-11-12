"""Ghi file .xls voi vo OLE2 tu dung (khong can file mau dong vai tro vo).

OLE2 v3, sector 512, stream 'Workbook' nam trong FAT thuong (khong dung mini
stream) de don gian va an toan. File tu dung se lon hon file mau goc nhung noi
dung/format lay tu khuon nen van mo dung bang Excel/LibreOffice.
"""
import struct

SECTOR = 512


def _dirent(name, typ, start, size, child=0xFFFFFFFF):
    n = name.encode("utf-16-le") + b"\x00\x00"
    body = n + b"\x00" * (64 - len(n))
    body += struct.pack("<HBB", len(n), typ, 1)
    body += struct.pack("<III", 0xFFFFFFFF, 0xFFFFFFFF, child)
    body += b"\x00" * 36
    body += struct.pack("<II", start, size)
    body += b"\x00" * 4
    return body


def write_ole_workbook(path, stream):
    """Dong goi stream BIFF8 thanh file .xls."""
    data = stream + b"\x00" * ((SECTOR - len(stream) % SECTOR) % SECTOR)
    n_data = len(data) // SECTOR
    # bo tri: [data sectors][FAT sectors][directory sector]
    max_sect = 128                         # 1 sector FAT = 128 muc
    n_fat = max(1, (n_data + 2 + max_sect - 1) // max_sect)
    fat_start = n_data
    dir_start = n_data + n_fat
    total_sectors = dir_start + 1
    if total_sectors > 128 * n_fat:
        n_fat = (total_sectors + 127) // 128
        fat_start = n_data
        dir_start = n_data + n_fat
        total_sectors = dir_start + 1

    fat = [0xFFFFFFFF] * (n_fat * 128)
    for i in range(n_data):
        fat[i] = i + 1 if i < n_data - 1 else 0xFFFFFFFE
    for i in range(n_fat):
        fat[fat_start + i] = 0xFFFFFFFE
    fat[dir_start] = 0xFFFFFFFE

    directory = _dirent("Root Entry", 5, 0xFFFFFFFE, 0, child=1)
    directory += _dirent("Workbook", 2, 0, len(stream))
    directory += b"\x00" * ((SECTOR - len(directory) % SECTOR) % SECTOR)

    header = bytearray(SECTOR)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<HHHHH", header, 24, 0x003E, 0x0003, 0xFFFE, 9, 6)
    struct.pack_into("<I", header, 40, 0)
    struct.pack_into("<I", header, 44, n_fat)
    struct.pack_into("<I", header, 48, dir_start)
    struct.pack_into("<I", header, 52, 0)
    struct.pack_into("<I", header, 56, 4096)
    struct.pack_into("<I", header, 60, 0xFFFFFFFE)
    struct.pack_into("<I", header, 64, 0)
    struct.pack_into("<I", header, 68, 0xFFFFFFFE)
    struct.pack_into("<I", header, 72, 0)
    for i in range(109):
        v = (fat_start + i) if i < n_fat else 0xFFFFFFFF
        struct.pack_into("<I", header, 76 + i * 4, v)

    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
    # sector FAT phai du so sector da khai bao
    fat_bytes = fat_bytes[:SECTOR * n_fat]
    while len(fat_bytes) < SECTOR * n_fat:
        fat_bytes += b"\xff" * SECTOR
    out = bytes(header) + data + fat_bytes + directory
    with open(path, "wb") as fh:
        fh.write(out)
    return path
