#!/usr/bin/env python3
"""
SL & OTIF generator — CLI wrapper around sl_otif_core.

Usage:
  python generate_sl_otif.py <file.xlsx> [more.xlsx ...] [-o OUTDIR] [-n 7]
  python generate_sl_otif.py            # all .xlsx beside this script
"""
import argparse, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sl_otif_core import process_file, TOL


def main():
    ap = argparse.ArgumentParser(description="Generate SL & OTIF result workbook(s) with live formulas.")
    ap.add_argument("files", nargs="*")
    ap.add_argument("-o", "--outdir", default="hasil")
    ap.add_argument("-n", "--tol", type=int, default=TOL)
    args = ap.parse_args()
    here = os.path.dirname(os.path.abspath(__file__))
    files = args.files or [os.path.join(here, f) for f in os.listdir(here)
                           if f.lower().endswith(".xlsx") and not f.startswith("~$")
                           and "SL & OTIF" not in f]
    outdir = args.outdir if os.path.isabs(args.outdir) else os.path.join(here, args.outdir)
    os.makedirs(outdir, exist_ok=True)
    for f in files:
        dst = os.path.join(outdir, "")
        dst, st, name = process_file(f, os.path.join(outdir, _outname(f)), args.tol)
        print(f"[OK] {os.path.basename(f)} -> {os.path.basename(dst)}  lines={st['nlines']} "
              f"months={st['order']}  FULFIL={st['fulfil']:.2%} OTIF={st['otif']:.2%}")
    print("Output:", outdir)


def _outname(path):
    import re
    b = re.sub(r"\.xlsx?$", "", os.path.basename(path), flags=re.I)
    b = re.sub(r"[^0-9A-Za-z ]+", " ", b).strip()
    return f"2026_{b.upper()}_SL & OTIF.xlsx"


if __name__ == "__main__":
    main()
