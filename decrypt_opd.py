"""Offline recovery, independent of Streamlit and GitHub credentials.

Install: python -m pip install cryptography
Run:     python decrypt_opd.py OPD_2026-08-03.xlsx.enc recovered_OPD.xlsx
The key is requested privately, rather than placed on the command line.
"""
import argparse
import getpass
import io
import sys
import zipfile
from pathlib import Path
from cryptography.fernet import Fernet, InvalidToken


def main() -> int:
    parser = argparse.ArgumentParser(description="Decrypt an OPD archived by the supplied app.")
    parser.add_argument("encrypted_file", type=Path)
    parser.add_argument("output_xlsx", type=Path)
    args = parser.parse_args()
    if args.output_xlsx.suffix.lower() != ".xlsx":
        parser.error("Choose an output filename ending in .xlsx")
    if args.output_xlsx.exists():
        parser.error("The output already exists. Choose a different filename to avoid overwriting it.")
    try:
        token = args.encrypted_file.read_bytes()
        key = getpass.getpass("Fernet encryption key: ").strip()
        raw = Fernet(key.encode("ascii")).decrypt(token)
        with zipfile.ZipFile(io.BytesIO(raw)) as workbook:
            if "xl/workbook.xml" not in workbook.namelist() or workbook.testzip() is not None:
                raise ValueError("Not a valid XLSX workbook")
        with args.output_xlsx.open("xb") as output:
            output.write(raw)
    except (OSError, ValueError, InvalidToken, zipfile.BadZipFile, UnicodeError):
        print("Recovery failed. Check the file, encryption key and output path; no valid recovery was confirmed.", file=sys.stderr)
        return 1
    print(f"Original workbook restored to {args.output_xlsx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
