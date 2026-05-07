"""
PDF Corruption Checker (Recursive)
------------------------------------
Scans all PDF files in a specified folder AND all its subfolders (root),
then reports which files are valid or corrupt with a detailed summary.

Usage:
    python check_pdf_corruption.py <folder_path>
    python check_pdf_corruption.py               # prompts for folder path

Requirements:
    pip install pypdf
"""

import os
import sys
from pathlib import Path
from datetime import datetime

# ── Install pypdf if missing ──────────────────────────────────────────────────
try:
    from pypdf import PdfReader
    from pypdf.errors import PdfReadError, PdfStreamError
except ImportError:
    import subprocess
    print("Installing required library: pypdf ...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pypdf", "-q"])
    from pypdf import PdfReader
    from pypdf.errors import PdfReadError, PdfStreamError


# ── ANSI colours (gracefully disabled on Windows if not supported) ────────────
def _supports_colour():
    return hasattr(sys.stdout, "isatty") and sys.stdout.isatty()

GREEN  = "\033[92m" if _supports_colour() else ""
RED    = "\033[91m" if _supports_colour() else ""
YELLOW = "\033[93m" if _supports_colour() else ""
CYAN   = "\033[96m" if _supports_colour() else ""
BOLD   = "\033[1m"  if _supports_colour() else ""
RESET  = "\033[0m"  if _supports_colour() else ""


# ── PDF validation ────────────────────────────────────────────────────────────
def check_pdf(filepath: Path) -> dict:
    """
    Validate a single PDF file.

    Returns a dict:
        status  : "OK" | "CORRUPT" | "EMPTY" | "NOT_A_PDF"
        pages   : int  (0 if unreadable)
        reason  : str  (empty string if OK)
        size_kb : float
    """
    result = {
        "path"   : filepath,
        "status" : "OK",
        "pages"  : 0,
        "reason" : "",
        "size_kb": round(filepath.stat().st_size / 1024, 2),
    }

    # 1. Empty file check
    if filepath.stat().st_size == 0:
        result["status"] = "EMPTY"
        result["reason"] = "File is 0 bytes"
        return result

    # 2. Magic-byte check  (%PDF header)
    try:
        with open(filepath, "rb") as fh:
            header = fh.read(5)
        if not header.startswith(b"%PDF"):
            result["status"] = "NOT_A_PDF"
            result["reason"] = f"Invalid PDF header: {header[:8]}"
            return result
    except OSError as exc:
        result["status"] = "CORRUPT"
        result["reason"] = f"Cannot read file: {exc}"
        return result

    # 3. Full structural check via pypdf
    try:
        reader = PdfReader(str(filepath), strict=False)

        # Try accessing every page to catch partial corruption
        page_count = len(reader.pages)
        for i, page in enumerate(reader.pages):
            try:
                _ = page.extract_text()
            except Exception as page_exc:
                result["status"] = "CORRUPT"
                result["reason"] = f"Page {i + 1} unreadable: {page_exc}"
                result["pages"] = page_count
                return result

        result["pages"] = page_count

    except PdfReadError as exc:
        result["status"] = "CORRUPT"
        result["reason"] = f"PdfReadError: {exc}"
    except PdfStreamError as exc:
        result["status"] = "CORRUPT"
        result["reason"] = f"PdfStreamError: {exc}"
    except Exception as exc:
        result["status"] = "CORRUPT"
        result["reason"] = f"Unexpected error: {exc}"

    return result


# ── Recursive folder scan ─────────────────────────────────────────────────────
def scan_folder(root: Path) -> list:
    """Walk *root* recursively and check every .pdf file found."""
    pdf_files = sorted(root.rglob("*.pdf"))
    pdf_files += sorted(root.rglob("*.PDF"))   # also catch uppercase extension

    # De-duplicate (rglob on case-insensitive FS may return dupes)
    seen   = set()
    unique = []
    for p in pdf_files:
        key = p.resolve()
        if key not in seen:
            seen.add(key)
            unique.append(p)

    results = []
    total   = len(unique)

    if total == 0:
        print(f"\n{YELLOW}No PDF files found under: {root}{RESET}\n")
        return results

    print(f"\n{BOLD}{CYAN}Scanning {total} PDF file(s) under:{RESET} {root}\n")
    print(f"{'#':<5} {'Status':<12} {'Pages':>6}  {'Size (KB)':>10}  Path")
    print("-" * 100)

    for idx, filepath in enumerate(unique, 1):
        res = check_pdf(filepath)
        results.append(res)

        # Relative path for compact display
        try:
            display_path = filepath.relative_to(root)
        except ValueError:
            display_path = filepath

        # Depth indicator (indentation based on subfolder depth)
        depth  = len(display_path.parts) - 1
        indent = "  " * depth + ("- " if depth > 0 else "")

        if res["status"] == "OK":
            colour, icon = GREEN, "[OK]"
        elif res["status"] in ("EMPTY", "NOT_A_PDF"):
            colour, icon = YELLOW, "[WARN]"
        else:
            colour, icon = RED, "[CORRUPT]"

        status_str = f"{colour}{icon:<10}{RESET}"
        print(f"{idx:<5} {status_str} {res['pages']:>6}  {res['size_kb']:>10.2f}  {indent}{display_path}")

        if res["reason"]:
            print(f"            {RED}  -> {res['reason']}{RESET}")

    return results


# ── Summary report ────────────────────────────────────────────────────────────
def print_summary(results: list, root: Path) -> int:
    total    = len(results)
    ok_files = [r for r in results if r["status"] == "OK"]
    corrupt  = [r for r in results if r["status"] == "CORRUPT"]
    empty    = [r for r in results if r["status"] == "EMPTY"]
    not_pdf  = [r for r in results if r["status"] == "NOT_A_PDF"]

    print("\n" + "=" * 100)
    print(f"{BOLD}SUMMARY  --  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}{RESET}")
    print("=" * 100)
    print(f"  Root folder  : {root}")
    print(f"  Total PDFs   : {total}")
    print(f"  {GREEN}[OK]      Valid   : {len(ok_files)}{RESET}")
    print(f"  {RED}[CORRUPT] Corrupt : {len(corrupt)}{RESET}")
    print(f"  {YELLOW}[WARN]    Empty   : {len(empty)}{RESET}")
    print(f"  {YELLOW}[WARN]    Not PDF : {len(not_pdf)}{RESET}")

    if corrupt:
        print(f"\n{BOLD}{RED}Corrupt files:{RESET}")
        for r in corrupt:
            print(f"  * {r['path']}")
            print(f"    Reason: {r['reason']}")

    if empty:
        print(f"\n{BOLD}{YELLOW}Empty files:{RESET}")
        for r in empty:
            print(f"  * {r['path']}")

    if not_pdf:
        print(f"\n{BOLD}{YELLOW}Files with .pdf extension but invalid PDF content:{RESET}")
        for r in not_pdf:
            print(f"  * {r['path']}")

    print("=" * 100)

    # Exit code: 0 = all OK, 1 = issues found
    return 0 if (not corrupt and not empty and not not_pdf) else 1


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    if len(sys.argv) >= 2:
        folder = sys.argv[1]
    else:
        folder = input("Enter the folder path to scan: ").strip().strip('"').strip("'")

    root = Path(folder).expanduser().resolve()

    if not root.exists():
        print(f"{RED}Error: Path does not exist -> {root}{RESET}")
        sys.exit(2)

    if not root.is_dir():
        print(f"{RED}Error: Path is not a directory -> {root}{RESET}")
        sys.exit(2)

    results   = scan_folder(root)
    exit_code = print_summary(results, root)
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
