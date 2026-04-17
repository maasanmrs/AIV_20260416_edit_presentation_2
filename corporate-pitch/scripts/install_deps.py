#!/usr/bin/env python3
"""Check and install required Python packages for corporate-pitch skill."""
import subprocess, sys

REQUIRED = [
    "python-pptx",
    "google-genai",
    "Pillow",
    "lxml",
    "python-docx",
    "PyPDF2",
]

def check_and_install():
    missing = []
    for pkg in REQUIRED:
        import_name = {
            "python-pptx": "pptx",
            "google-genai": "google.genai",
            "Pillow": "PIL",
            "python-docx": "docx",
            "PyPDF2": "PyPDF2",
        }.get(pkg, pkg)
        try:
            __import__(import_name)
            print(f"  [OK] {pkg}")
        except ImportError:
            missing.append(pkg)
            print(f"  [--] {pkg} (not installed)")

    if not missing:
        print("\n[DEPS] All packages available.")
        return True

    print(f"\n[DEPS] Installing {len(missing)} missing package(s)...")
    for pkg in missing:
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", pkg, "-q"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
            )
            print(f"  [OK] {pkg} installed")
        except subprocess.CalledProcessError as e:
            print(f"  [FAIL] {pkg}: {e}")
            print(f"         Manual install: pip install {pkg}")
            return False
    print("\n[DEPS] All packages installed.")
    return True

if __name__ == "__main__":
    print("[DEPS] Checking dependencies...")
    ok = check_and_install()
    sys.exit(0 if ok else 1)
