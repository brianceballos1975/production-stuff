"""
install_fonts.py — Download & install Google Fonts needed for the current gavel batch.

Reads summary.csv to find which fonts were used, cross-references against the
known Google Fonts list, downloads .ttf files from the Google Fonts GitHub
repository, and installs them into the Windows user font directory
(%LOCALAPPDATA%\\Microsoft\\Windows\\Fonts) -- no admin rights needed.

CorelDRAW 2026 (and any app that reads the user font registry key) will see
these fonts immediately; a restart of CorelDRAW may be required.

Usage:
    python install_fonts.py                  # uses gavel_eps/summary.csv
    python install_fonts.py --csv path/to/summary.csv
    python install_fonts.py --fonts "Lato" "Playfair Display"
"""

import argparse
import csv
import ctypes
import io
import json
import os
import shutil
import sys
import urllib.request
import winreg
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── Known Google Fonts (lowercase) ────────────────────────────────────────────
# Keep in sync with gavel_eps_generator.py _GOOGLE_FONTS set.
GOOGLE_FONTS = {
    "lato", "open sans", "roboto", "roboto slab", "montserrat", "oswald",
    "raleway", "source sans pro", "ubuntu", "nunito", "poppins", "merriweather",
    "playfair display", "old standard tt", "homemade apple", "dancing script",
    "great vibes", "pacifico", "lobster", "abril fatface", "noto serif",
    "noto sans", "pt sans", "pt serif", "libre baskerville", "cormorant garamond",
}

# GitHub raw base for downloading fonts
GITHUB_API  = "https://api.github.com/repos/google/fonts/contents"
GITHUB_RAW  = "https://raw.githubusercontent.com/google/fonts/main"
# Subdirectories to search in order (most fonts are in ofl/, some in apache/)
FONT_BASES  = ["ofl", "apache", "ufl"]

# User font directory — no admin required on Windows 10 1809+
USER_FONTS_DIR = Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Windows" / "Fonts"
USER_FONT_REG  = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"


# ── GitHub helpers ─────────────────────────────────────────────────────────────

def _gh_get(path: str) -> object:
    """GET a GitHub API endpoint and return parsed JSON."""
    req = urllib.request.Request(
        f"{GITHUB_API}/{path}",
        headers={"User-Agent": "gavel-font-installer", "Accept": "application/vnd.github+json"},
    )
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read())


def _family_to_dir(family: str) -> str:
    """Convert a font family name to its GitHub directory name (lowercase, no spaces)."""
    return family.lower().replace(" ", "")


def find_font_dir(family: str) -> tuple[str, str] | None:
    """
    Locate the font in the Google Fonts GitHub repo.
    Returns (base, dirname) e.g. ("ofl", "lato"), or None if not found.
    """
    dirname = _family_to_dir(family)
    for base in FONT_BASES:
        try:
            _gh_get(f"{base}/{dirname}")   # just checking 200 vs 404
            return base, dirname
        except Exception:
            continue
    return None


def list_font_ttfs(base: str, dirname: str) -> list[dict]:
    """
    Return GitHub file-info dicts for all .ttf files in a font directory.
    Also checks a 'static/' subdirectory (used by variable-font families that
    also ship static instances for app compatibility).
    """
    entries = _gh_get(f"{base}/{dirname}")
    ttfs    = [e for e in entries if e["name"].lower().endswith(".ttf")]

    # If the directory has a static/ sub-folder, prefer those static .ttf files
    # (variable .ttf files won't install cleanly as named instances)
    static_dirs = [e for e in entries if e["type"] == "dir" and e["name"].lower() == "static"]
    if static_dirs:
        static_entries = _gh_get(f"{base}/{dirname}/static")
        static_ttfs    = [e for e in static_entries if e["name"].lower().endswith(".ttf")]
        if static_ttfs:
            ttfs = static_ttfs   # prefer static over variable

    return ttfs


# ── Font installation ──────────────────────────────────────────────────────────

def download_bytes(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "gavel-font-installer"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def install_font_file(src_bytes: bytes, filename: str) -> bool:
    """
    Write font bytes to the user fonts dir and register in the user registry.
    Returns True if newly installed, False if already present.
    """
    USER_FONTS_DIR.mkdir(parents=True, exist_ok=True)
    dst = USER_FONTS_DIR / filename

    if dst.exists():
        return False   # already installed

    dst.write_bytes(src_bytes)

    # Register in user-level font registry so all apps pick it up
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, USER_FONT_REG, 0, winreg.KEY_SET_VALUE
        ) as key:
            reg_name = Path(filename).stem + " (TrueType)"
            winreg.SetValueEx(key, reg_name, 0, winreg.REG_SZ, str(dst))
    except OSError:
        pass   # non-critical; file is already on disk

    # Tell GDI to load the new font so running apps see it without restart
    try:
        ctypes.windll.gdi32.AddFontResourceExW(str(dst), 0x10, 0)
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001D, 0, 0, 0x0002, 1000, None)
    except Exception:
        pass

    return True


def install_family(family: str) -> tuple[int, int]:
    """
    Find, download, and install all .ttf files for one font family.
    Returns (newly_installed, already_present) counts.
    """
    print(f"\n[{family}]")

    location = find_font_dir(family)
    if not location:
        print(f"  ERROR: '{family}' not found in Google Fonts GitHub repo.")
        return 0, 0

    base, dirname = location
    print(f"  Found at: {base}/{dirname}")

    ttf_entries = list_font_ttfs(base, dirname)
    if not ttf_entries:
        print(f"  No .ttf files found.")
        return 0, 0

    new_count = already_count = 0
    for entry in ttf_entries:
        filename = entry["name"]
        raw_url  = entry.get("download_url") or f"{GITHUB_RAW}/{base}/{dirname}/{filename}"
        print(f"  Downloading {filename} …", end="", flush=True)
        try:
            font_bytes = download_bytes(raw_url)
            if install_font_file(font_bytes, filename):
                print(" installed ✓")
                new_count += 1
            else:
                print(" already installed")
                already_count += 1
        except Exception as e:
            print(f" ERROR: {e}")

    return new_count, already_count


# ── CSV helpers ────────────────────────────────────────────────────────────────

def fonts_from_csv(csv_path: Path) -> list[str]:
    """Return sorted list of unique font names in summary.csv."""
    fonts: set[str] = set()
    with open(csv_path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            fnt = row.get("font", "").strip()
            if fnt:
                fonts.add(fnt)
    return sorted(fonts)


# ── main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Install Google Fonts for gavel batch")
    parser.add_argument("--csv",   default="gavel_eps/summary.csv",
                        help="Path to summary.csv (default: gavel_eps/summary.csv)")
    parser.add_argument("--fonts", nargs="*",
                        help="Install specific font family names instead of reading CSV")
    args = parser.parse_args()

    print("=" * 60)
    print("Google Font Installer for Gavel Bands")
    print("=" * 60)

    if args.fonts:
        all_fonts = args.fonts
    else:
        csv_path = Path(args.csv)
        if not csv_path.exists():
            raise SystemExit(f"summary.csv not found at {csv_path}")
        all_fonts = fonts_from_csv(csv_path)
        print(f"\nFonts used in this batch ({len(all_fonts)}):")
        for f in all_fonts:
            tag = "(Google Font)" if f.lower() in GOOGLE_FONTS else "(system font - skipping)"
            print(f"  {f:30s}  {tag}")

    google_needed = [f for f in all_fonts if f.lower() in GOOGLE_FONTS]

    if not google_needed:
        print("\nNo Google Fonts needed - all are Windows system fonts.")
        return

    print(f"\nInstalling {len(google_needed)} Google Font family/families from GitHub ...")

    total_new = total_already = 0
    for family in google_needed:
        new, already = install_family(family)
        total_new    += new
        total_already += already

    print(f"\n{'='*60}")
    print(f"New files installed : {total_new}")
    print(f"Already present     : {total_already}")
    if total_new:
        print(f"Install location    : {USER_FONTS_DIR}")
        print("\nRestart CorelDRAW (if open) to pick up the new fonts.")
    print("=" * 60)


if __name__ == "__main__":
    main()
