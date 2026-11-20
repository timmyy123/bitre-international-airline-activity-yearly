"""
Download the BITRE International Airline Activity monthly tables listed in source_monthly/manifest.txt.

Each manifest line is: YYYY-MM <url> [optional note]
Files are saved as source_monthly/YYYY-MM.xlsx (or .xls). Existing files are skipped.

To add a new month: append a line to manifest.txt with the link from
https://www.bitre.gov.au/resource/aviation/international-airline-activity-monthly-publications
then run this script and build_yearly.py.
"""
import os
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "source_monthly")


def main():
    with open(os.path.join(SRC, "manifest.txt")) as f:
        entries = [line.split()[:2] for line in f if line.strip() and not line.startswith("#")]
    failed = 0
    for ym, url in entries:
        ext = os.path.splitext(url)[1].lower()
        dest = os.path.join(SRC, f"{ym}{ext}")
        if os.path.exists(dest) and os.path.getsize(dest) > 0:
            continue
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=60) as r, open(dest, "wb") as out:
                out.write(r.read())
            print("downloaded", ym)
        except Exception as e:  # keep going; report at the end
            failed += 1
            print("FAILED", ym, url, e, file=sys.stderr)
    print(f"{len(entries)} months in manifest, {failed} failed")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
