#!/usr/bin/env python3
"""Convert a meeting transcript to plain UTF-8 text for reading and verification.

    python3 to_text.py <transcript.(rtf|docx|txt|vtt)> [--outdir DIR]

Writes <outdir>/<name>.txt (outdir defaults to a fresh scratch directory) and prints JSON
with the path, the method used and the line count. Converters, in order:
  .rtf / .docx  LibreOffice if installed (any platform), else textutil (built into macOS),
                else exit 1 with an install hint. pandoc does not read these RTFs.
  .vtt          handled here: WEBVTT header, cue identifiers, timestamp lines and inline
                tags stripped; <v Speaker> voice tags become "Speaker: "; rolling-caption
                repeats collapsed.
  .txt          copied through.
"""
import argparse, json, os, re, shutil, subprocess, sys, tempfile
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_report import find_soffice, soffice_run

TIMESTAMP = re.compile(r"^\s*(\d{1,2}:)?\d{2}:\d{2}[.,]\d{3}\s*-->")
VOICE = re.compile(r"<v\s+([^>]+)>")
TAG = re.compile(r"<[^>]+>")
SKIP_PREFIX = ("WEBVTT", "NOTE", "STYLE", "REGION")


def via_soffice(src, outdir):
    exe = find_soffice()
    if not exe:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        try:
            soffice_run(exe, ["--convert-to", "txt:Text (encoded):UTF8", "--outdir", outdir, src], tmp)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
            return None
    out = os.path.join(outdir, Path(src).stem + ".txt")
    return out if os.path.exists(out) else None


def via_textutil(src, outdir):
    if not shutil.which("textutil"):
        return None
    # -stdout: textutil then needs no write access of its own, which matters in sandboxes.
    try:
        r = subprocess.run(["textutil", "-convert", "txt", "-encoding", "UTF-8", "-stdout", src],
                           check=True, capture_output=True, timeout=120)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        return None
    out = os.path.join(outdir, Path(src).stem + ".txt")
    with open(out, "wb") as f:
        f.write(r.stdout)
    return out


def vtt_to_text(src, outdir):
    lines = open(src, encoding="utf-8-sig", errors="replace").read().splitlines()
    out_lines, prev = [], None
    for i, line in enumerate(lines):
        st = line.strip()
        nxt = lines[i + 1].strip() if i + 1 < len(lines) else ""
        # A cue identifier is whatever line sits directly above a timestamp line.
        if not st or st.startswith(SKIP_PREFIX) or TIMESTAMP.match(st) or TIMESTAMP.match(nxt):
            continue
        st = VOICE.sub(lambda m: m.group(1).strip() + ": ", st)
        st = re.sub(r"\s+", " ", TAG.sub("", st)).strip()
        if st and st != prev:
            out_lines.append(st)
            prev = st
    out = os.path.join(outdir, Path(src).stem + ".txt")
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(out_lines) + "\n")
    return out


def main():
    ap = argparse.ArgumentParser(description="Convert a transcript to plain text.")
    ap.add_argument("transcript")
    ap.add_argument("--outdir", help="where to write the .txt (default: a new scratch directory)")
    a = ap.parse_args()

    src = os.path.abspath(a.transcript)
    if not os.path.exists(src):
        sys.exit(f"Not found: {src}")
    outdir = os.path.abspath(a.outdir) if a.outdir else tempfile.mkdtemp(prefix="meeting-report-")
    os.makedirs(outdir, exist_ok=True)
    ext = Path(src).suffix.lower()

    if ext == ".txt":
        out = os.path.join(outdir, Path(src).name)
        if out != src:
            shutil.copy(src, out)
        method = "copy"
    elif ext == ".vtt":
        out, method = vtt_to_text(src, outdir), "vtt"
    elif ext in (".rtf", ".docx", ".doc", ".odt"):
        out, method = via_soffice(src, outdir), "libreoffice"
        if not out:
            out, method = via_textutil(src, outdir), "textutil"
        if not out:
            sys.exit(f"Cannot convert {ext} here: LibreOffice is not installed and textutil is "
                     "unavailable. Install LibreOffice (or set SOFFICE=/path/to/soffice) and retry.")
    else:
        sys.exit(f"Unsupported transcript type {ext!r}. Use .rtf, .docx, .txt or .vtt.")

    text = open(out, encoding="utf-8", errors="replace").read()
    if not text.strip():
        sys.exit(f"Conversion produced an empty file: {out}")
    n = text.count("\n") + (0 if text.endswith("\n") else 1)
    print(json.dumps({"output": out, "method": method, "lines": n, "characters": len(text),
                      "note": "over 2000 lines - read it in offsets to the last line" if n > 2000 else ""},
                     indent=2))


if __name__ == "__main__":
    main()
