#!/usr/bin/env python
#  midnam2xlsx.py
#
#  Copyright 2020 Christopher Arndt <chris@chrisarndt.de>
#
"""Convert MIDNAM instrument description into a patch list excel spread sheet."""

import argparse
import logging
import sys

from lxml import etree
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from .styles import colhead, numcell, txtcell


log = logging.getLogger(__file__)

COLUMNS_ALL = {
    "A": ("Bank", 8, colhead, numcell),
    "B": ("MSB", 7, colhead, numcell),
    "C": ("LSB", 7, colhead, numcell),
    "D": ("PC", 7, colhead, numcell),
    "E": ("Category", 10, colhead, numcell),
    "F": ("Name", 20, colhead, txtcell),
}

COLUMNS_BANK = {
    "A": ("PC", 6, colhead, numcell),
    "B": ("Category", 8, colhead, numcell),
    "C": ("Name", 20, colhead, txtcell),
}


def sanitize(s):
    for repl in r"\/*?:[]":
        s = s.replace(repl, " ")
    return s


def parse_midnam(fileobj):
    tree = etree.parse(fileobj)
    namesets = {}

    for nsnum, nameset in enumerate(tree.xpath("*/ChannelNameSet")):
        nsname = nameset.get("Name", "<unnamed set #{}>".format(nsnum + 1))
        log.debug("Channel Name Set: %s", nsname)

        patchbanks = {}

        for pbnum, patchbank in enumerate(nameset.xpath("./PatchBank")):
            pbname = patchbank.get("Name", "<unnamed bank #{}>".format(pbnum + 1))
            log.debug("Patch Bank: %s", pbname)

            msb = patchbank.xpath("./MIDICommands/ControlChange[@Control='0']")

            try:
                msb = int(msb[0].get("Value"))
            except:
                msb = ""

            lsb = patchbank.xpath("./MIDICommands/ControlChange[@Control='32']")

            try:
                lsb = int(lsb[0].get("Value"))
            except:
                lsb = ""

            patches = []

            for pnum, patch in enumerate(patchbank.xpath("./PatchNameList/Patch")):
                pnumber = patch.get("Number")
                pname = patch.get(
                    "Name", "<unnamed patch #{}".format(patch.get("Number", pnum + 1))
                )

                # If name is the same as the patch number, skip this patch
                try:
                    if int(pnumber) == int(pname):
                        continue
                except (TypeError, ValueError):
                    pass

                try:
                    pc = int(patch.get("ProgramChange"))
                except:
                    pc = ""

                try:
                    category, pname = pname.split(":", 1)
                except:
                    category = ""

                patches.append([pbname, msb, lsb, pc, category, pname])

            if patches:
                patchbanks[pbname] = patches

        if patchbanks:
            namesets[nsname] = patchbanks

    return namesets


def write_xlsx(nsname, nameset):
    wb = Workbook()
    wb.add_named_style(colhead)

    # Use first work sheet named 'All' to collect all patches from all banks
    ws_all = wb.active
    ws_all.title = "All"
    ws_all_row = 1
    ws_all.freeze_panes = "A2"

    # Set column header styles for 'All' worksheet
    for col, (value, width, headstyle, _) in COLUMNS_ALL.items():
        ws_all[col + "1"] = value
        ws_all[col + "1"].style = headstyle
        ws_all.column_dimensions[col].width = width

    # Create one work sheet for each patch bank
    for pbnum, (pbname, patchbank) in enumerate(nameset.items()):
        ws = wb.create_sheet(title=sanitize(pbname))
        ws.freeze_panes = "A2"

        # Set column header styles for patch bank worksheet
        for col, (value, width, headstyle, _) in COLUMNS_BANK.items():
            ws[col + "1"] = value
            ws[col + "1"].style = headstyle
            ws.column_dimensions[col].width = width

        for row, patch in enumerate(patchbank, start=2):
            if pbname != "GM":
                ws_all.append(patch)
                ws_all_row += 1

                # Set cell styles for row in 'All' work sheet
                for column in COLUMNS_ALL:
                    ws_all["%s%i" % (column, ws_all_row)].style = COLUMNS_ALL[column][3]

            ws.append(patch[3:])

            # Set cell styles for row in patch bank work sheet
            for column in COLUMNS_BANK:
                ws["%s%i" % (column, row)].style = COLUMNS_BANK[column][3]

    # Set auto filter for 'All' work sheet
    ws_all.auto_filter.ref = ws_all.dimensions

    # Generate filename from nameset name
    basename = sanitize(nsname).strip()

    if basename:
        xlsxname = basename + ".xlsx"
    else:
        xlsxname = "nameset-%02i.xlsx" % (nsnum + 1)

    log.info("Writing '%s'...", xlsxname)
    wb.save(xlsxname)


def main(args=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("-v", "--verbose", action="store_true", help="Be verbose")
    ap.add_argument("midnam", type=open, help="MIDNAM input file")

    args = ap.parse_args(args)

    logging.basicConfig(
        format="%(name)s: %(levelname)s - %(message)s",
        level=logging.DEBUG if args.verbose else logging.INFO,
    )

    with args.midnam:
        namesets = parse_midnam(args.midnam)

    for nsname, nameset in namesets.items():
        write_xlsx(nsname, nameset)

    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
