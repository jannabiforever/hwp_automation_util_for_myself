import pythoncom
from hwp import *
import argparse
import pdfutil
import os
import osutil
import logging

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

pythoncom.CoInitialize()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        prog="hauto", description="Hwpx automation tool")
    sub_parsers = parser.add_subparsers(dest="command")

    converter = sub_parsers.add_parser(
        'convert', help="Convert hwp/hwpx file to other formats")
    converter.add_argument('--format', type=str, choices=['pdf', 'hwpx'])
    converter.add_argument('--path', type=str)
    converter.add_argument('-f', '--folder', action='store_true')
    converter.add_argument('--output', '-o', type=str)

    ns = parser.parse_args()

    try:
        app = HApp()
        if ns.command == 'convert':
            if ns.format == 'pdf':
                svb = ConvertToPdfServiceBuilder()
                svb.set_file_path(ns.path)\
                    .set_save_path(ns.output)\
                    .build()\
                    .execute(app)

            if ns.format == "hwpx":
                ConvertToHwpxServiceBuilder()\
                    .set_file_path(ns.path if not ns.folder else osutil.get_hwp_and_hwpx_files_under(ns.path))\
                    .set_save_path(ns.output)\
                    .build()\
                    .execute(app)

    finally:
        app.Quit()
