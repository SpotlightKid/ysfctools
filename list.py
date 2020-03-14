#!/usr/bin/env python

import argparse
import struct
import sys


FILE_MAGIC = "YAMAHA-YSFC"


def read_catalog(inputfile, types=None):
    data = inputfile.read(64)

    if len(data) != 64:
        raise ValueError("Invalid file header size.")

    magic = data[:16].decode().rstrip("\0")

    if magic != FILE_MAGIC:
        raise ValueError("Invalid file header magic string.")

    try:
        version = tuple(int(x) for x in data[16:32].rstrip(b"\0").split(b"."))
        assert len(version) == 3
    except:
        raise ValueError("Invalid file version format.")

    if version[0] < 1 or version[1] != 0 or version[2] not in (0, 1, 2, 3):
        raise ValueError("Unsupported file format version.")

    if version[0] >= 4:
        pad_size = struct.unpack(">I", data[48:52])[0]
        #inputfile.seek(pad_size, 1)
    elif data[36:] != 28 * b"\xff":
        raise ValueError("Invalid header padding.")

    catalog_size = struct.unpack(">I", data[32:36])[0]
    catalog_offset = inputfile.tell()
    catalog = []

    while inputfile.tell() < catalog_offset + catalog_size:
        block_id = inputfile.read(4).decode()

        if len(block_id) != 4:
            raise ValueError("Truncated catalogue.")
        elif not block_id.isalpha() or not block_id.isupper():
            raise ValueError("Invalid block identifier '%s' in catalogue." % block_id)

        try:
            offset = struct.unpack(">I", inputfile.read(4))[0]
        except:
            raise ValueError("Invalid catalogue entry '%s'. Could not read offset" % (block_id,))

        catalog.append((block_id, offset))

    blocks = {}

    for cblock_id, offset in catalog:
        if not cblock_id.startswith('E'):
            continue

        inputfile.seek(offset)
        block_id = inputfile.read(4).decode()

        if len(block_id) != 4:
            raise ValueError("Truncated block header. Expected '%s', read '%s'." %
                             (cblock_id, block_id))
        elif not block_id.isalpha() or not block_id.isupper():
            raise ValueError("Invalid block identifier '%s' in catalogue." % block_id)
        elif cblock_id != block_id:
            print("Wrong block ID at offset %s. Expected '%s', read '%s'." % (cblock_id, block_id))
            print("Ignoring block.")
            continue

        try:
            size = struct.unpack(">I", inputfile.read(4))[0]
        except:
            raise ValueError("Invalid block '%s'. Could not read size" % (block_id,))

        if not types or block_id in types:
            block_data = inputfile.read(size)

            if len(block_data) != size:
                raise ValueError("Truncated block '%s'. Expected %d bytes, got %d." %
                                 (block_id, size, len(block_data)))

            blocks[block_id] = block_data

    return version, blocks


def parse_entry_list(version, data):
    cursor = 4
    items = []
    count = struct.unpack_from(">I", data)[0]

    while cursor < len(data):
        magic, length = struct.unpack_from(">4sI", data, cursor)

        if magic != b"Entr" or cursor + length + 8 > len(data):
            raise ValueError("Invalid entry list block.")

        if version <= (1, 0, 2):
            size, offset, number = struct.unpack_from(">4xI4x2I", data, cursor + 8)
        else:
            size, offset, number = struct.unpack_from(">3I", data, cursor + 8)

        if version <= (1, 0, 1):
            names = data[cursor + 29:cursor + length + 8]
        elif version <= (1, 0, 2):
            names = data[cursor + 30:cursor + length + 8]
        else:
            names = data[cursor + 20:cursor + length + 8]

        names = names.strip(b"\0").split(b"\0")

        item = {
            "size": size,
            "offset": offset,
            "number": number,
            "name": names[0].decode().rstrip(),
        }

        if len(names) > 1:
            item["filename"] = names[1].decode().rstrip()

        if len(names) > 2:
            item["depends"] = names[2:]

        items.append(item)
        cursor += length + 8

    if len(items) != count:
        raise ValueError("Invalid file format")
    else:
        return items


def bankname(number):
    bank, program = number >> 8, 1 + (number & 0xFF)

    if bank in range(0x3F08, 0x3F10):
        return "USR%d:%03d" % (bank - 0x3F07, program)

    if bank == 0x3F28:
        return "USRDR:%03d" % program

    if bank in range(0x3F80, 0x3FC0):
        if program <= 128:
            return "SNG%d:SP%03d" % (bank - 0x3F7F, program)
        else:
            return "SNG%d:MV%03d" % (bank - 0x3F7F, program - 128)

    if bank in range(0x3FC0, 0x4000):
        if program <= 128:
            return "PTN%d:SP%03d" % (bank - 0x3FBF, program)
        else:
            return "PTN%d:MV%03d" % (bank - 0x3FBF, program - 128)

    return "0x%06x" % number


def main(args=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("inputfile", help="Input filename")

    args = ap.parse_args(args)

    with open(args.inputfile, 'rb') as fp:
        version, blocks = read_catalog(fp)

    print("Version:", version)

    for block in blocks:
        if not block.startswith("E"):
            continue

        for item in parse_entry_list(version, blocks[block]):
            if block == "EVCE":
                print("VCE %s %s" % (bankname(item["number"]), item["name"]))
            else:
                print("%3s %04d %s" % (block[1:], item["number"] + 1, item["name"]))


if __name__ == "__main__":
    sys.exit(main() or 0)
