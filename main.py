import pythoncom
from hwp import *
import argparse
import pdfutil
import logging

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

pythoncom.CoInitialize()


if __name__ == "__main__":
    app = HApp()

    parser = argparse.ArgumentParser(
        prog="hauto", description="Hwpx automation tool")
    sub_parsers = parser.add_subparsers(dest="command")

    converter = sub_parsers.add_parser(
        'convert', help="Convert hwp/hwpx file to other formats")
    converter.add_argument('--format', type=str, choices=['pdf'])
    converter.add_argument('--file', type=str, required=True)
    converter.add_argument('--output', '-o', type=str)

    ns = parser.parse_args()

    try:
        if ns.command == 'convert':
            if ns.format == 'pdf':
                svb = ConvertToPdfServiceBuilder()
                svb.set_file_path(ns.file)\
                    .set_save_path(ns.output)\
                    .build()\
                    .execute(app)
    finally:
        app.Quit()
