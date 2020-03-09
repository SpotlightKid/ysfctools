#!/usr/bin/env python

import argparse
import struct
import sys


FILE_MAGIC = "YAMAHA-YSFC"


def read_catalogue(inputfile, types=None):
    data = inputfile.read(64)

    if len(data) != 64:
        raise ValueError("Invalid file header size.")

    magic = data[:16].decode().rstrip("\0")

    if magic != FILE_MAGIC:
        raise ValueError("Invalid file header magic string.")

    if data[36:] != 28 * b"\xff":
        raise ValueError("Invalid header padding.")

    try:
        version = tuple(int(x) for x in data[16:32].rstrip(b"\x00").split(b"."))
        size = struct.unpack(">I", data[32:36])[0]
        catalog = inputfile.read(size)
        assert len(version) == 3 and len(catalog) == size
    except:
        raise ValueError("Truncated file")

    if version[0:2] != (1, 0) or version[2] not in (0, 1, 2, 3):
        raise ValueError("Unsupported file format version")

    blocks = {}
    cursor = 64

    while True:
        block_id = inputfile.read(4).decode()

        if not block_id:
            break

        if len(block_id) != 4:
            raise ValueError("Truncated file")
        elif not block_id.isalpha() or not block_id.isupper():
            raise ValueError("Invalid block identifier '%s' in catalogue." % block_id)

        try:
            size = struct.unpack(">I", inputfile.read(4))[0]
        except:
            raise ValueError("Invalid block '%s'. Could not read size" % (block_id,))

        cursor += 8 + size

        if not types or block_id in types:
            block_data = inputfile.read(size)

            if len(block_data) != size:
                raise ValueError("Truncated block '%s'. Expected %d bytes, got %d." %
                                 (block_id, size, len(block_data)))

            blocks[block_id] = block_data
        else:
            inputfile.seek(size, 1)
            pos = inputfile.tell()

            if pos != cursor:
                raise ValueError("Truncated block '%s'. EOF at %d bytes." % (block_id, pos))

    return version, blocks


def parse_entry_list(version, entries, data):
    cursor = 4
    items = []
    count = struct.unpack_from(">I", entries)[0]

    while cursor < len(entries):
        magic, length = struct.unpack_from(">4sI", entries, cursor)

        if magic != b"Entr" or cursor + length + 8 > len(entries):
            raise ValueError("Invalid file format")

        if version <= (1, 0, 2):
            size, offset, number = struct.unpack_from(">4xI4x2I", entries, cursor + 8)
        else:
            size, offset, number = struct.unpack_from(">3I", entries, cursor + 8)

        if version <= (1, 0, 1):
            names = entries[cursor + 29:cursor + length + 8]
        elif version <= (1, 0, 2):
            names = entries[cursor + 30:cursor + length + 8]
        else:
            names = entries[cursor + 20:cursor + length + 8]

        names = names.strip(b"\0").split(b"\0")

        item = {
            "number": number,
            "name": names[0].decode().rstrip(),
        }

        if len(names) > 1:
            item["filename"] = names[1].decode().rstrip()

        if len(names) > 2:
            item["depends"] = names[2:]
        if data:
            item["data"] = data[offset - 8 : offset + size]

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
        version, blocks = read_catalogue(fp)

    print("Version:", version)
    for block in blocks:
        size = struct.unpack_from(">I", blocks[block])[0]
        print("Block: %s %4d %r" % (block, size, blocks[block][4:20]))

    for block in blocks:
        if not block.startswith("E"):
            continue

        for item in parse_entry_list(version, blocks[block], None):
            if block == "EVCE":
                print("VCE %s %s" % (bankname(item["number"]), item["name"]))
            else:
                print("%3s %04d %s" % (block[1:], item["number"] + 1, item["name"]))


if __name__ == "__main__":
    sys.exit(main() or 0)
