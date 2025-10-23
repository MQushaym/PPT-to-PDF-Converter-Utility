#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Convert PPT/PPTX files to PDF using LibreOffice (headless), no PowerPoint, no PowerShell.
- Works on Windows/macOS/Linux (as long as LibreOffice is installed and `soffice` is accessible)
- Default behavior: convert only files in the given folder (no recursion)
- Optional: --recursive to include subfolders
Usage examples:
    python ppt_to_pdf_lo.py "C:\path\to\folder"
    python ppt_to_pdf_lo.py "C:\path\to\folder" --recursive
    python ppt_to_pdf_lo.py "C:\path\to\folder" --soffice "C:\Program Files\LibreOffice\program\soffice.exe"
"""

import argparse
import shutil
import subprocess
from pathlib import Path
import sys
import time

def find_soffice(user_path: str | None) -> str:
    # 1) explicit path from user
    if user_path:
        p = Path(user_path)
        if p.exists():
            return str(p)
        print(f"[!] Provided soffice path not found: {user_path}", file=sys.stderr)
    # 2) common Windows path
    common_win = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
    if common_win.exists():
        return str(common_win)
    # 3) PATH lookup
    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if found:
        return found
    # 4) give up
    raise FileNotFoundError(
        "LibreOffice 'soffice' executable not found. Install LibreOffice or pass --soffice PATH"
    )

def collect_files(root: Path, recursive: bool) -> list[Path]:
    exts = (".ppt", ".pptx")
    if recursive:
        return [p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in exts]
    else:
        return [p for p in root.glob("*") if p.is_file() and p.suffix.lower() in exts]

def convert_file(soffice: str, fpath: Path) -> tuple[bool, str]:
    outdir = str(fpath.parent)
    cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, str(fpath)]
    try:
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        # LibreOffice prints info on stdout; we consider success if PDF exists
        out_pdf = fpath.with_suffix(".pdf")
        ok = out_pdf.exists()
        msg = res.stdout.strip() or res.stderr.strip()
        return ok, msg
    except Exception as e:
        return False, str(e)

def main():
    ap = argparse.ArgumentParser(
        description="Batch convert PPT/PPTX to PDF via LibreOffice (headless)."
    )
    ap.add_argument("folder", type=str, help="Folder containing PPT/PPTX files")
    ap.add_argument("--recursive", action="store_true", help="Include subfolders")
    ap.add_argument("--soffice", type=str, default=None, help="Path to soffice executable")
    args = ap.parse_args()

    root = Path(args.folder).expanduser().resolve()
    if not root.exists() or not root.is_dir():
        print(f"[x] Folder not found or not a directory: {root}", file=sys.stderr)
        sys.exit(2)

    try:
        soffice = find_soffice(args.soffice)
    except FileNotFoundError as e:
        print(f"[x] {e}", file=sys.stderr)
        sys.exit(3)

    files = collect_files(root, args.recursive)
    total = len(files)
    print(f"🔎 Found {total} PPT/PPTX file(s) in: {root} (recursive={args.recursive})")
    if total == 0:
        sys.exit(0)

    ok = 0
    fail = 0
    t0 = time.time()
    for i, f in enumerate(files, 1):
        print(f"[{i}/{total}] Converting: {f.name} ...", end="", flush=True)
        success, info = convert_file(soffice, f)
        if success:
            ok += 1
            print(" ✅")
        else:
            fail += 1
            print(" ❌")
            print(f"    ↳ {info}")

    dt = time.time() - t0
    print("=" * 48)
    print(f"Done. Success={ok} | Failed={fail} | Total={total} | Time={dt:.1f}s")

if __name__ == "__main__":
    main()
