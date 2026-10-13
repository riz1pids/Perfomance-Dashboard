#!/usr/bin/env python3
"""Bring the months already on the published site into the build folder.

A GitHub Pages build starts from an empty folder, so without this step every
run would publish only the month it just built. This copies each published
month into OUTPUT_DIR first; build_dashboard_data.py then replaces this
month, keeps the others and prunes beyond KEEP_MONTHS.

A month is copied only if every one of its files downloads; a partial month
is skipped rather than published broken. No published site yet is not an
error - the first run simply starts empty.

Usage: python ci/fetch_published.py <site url> <output data dir>
"""
import json
import os
import shutil
import sys
import tempfile
import urllib.parse
import urllib.request


def get(url, timeout=120):
    with urllib.request.urlopen(url, timeout=timeout) as r:
        return r.read()


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    site, out = sys.argv[1].rstrip("/"), sys.argv[2]
    try:
        man = json.loads(get(site + "/data/manifest.json", timeout=30))
    except Exception as e:                      # 404 on the first run
        print(f"No published months to carry over ({e.__class__.__name__}: {e}).")
        return
    months = man.get("months", [])
    os.makedirs(out, exist_ok=True)
    kept = []
    for m in months:
        mid, files = m.get("id", ""), m.get("files", [])
        if not mid or not files:
            continue
        tmp = tempfile.mkdtemp(prefix="month-")
        try:
            for f in files:
                url = site + "/data/" + urllib.parse.quote(f)
                dst = os.path.join(tmp, f)
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                with open(dst, "wb") as fh:
                    fh.write(get(url))
        except Exception as e:
            print(f"  skipped {m.get('label', mid)}: {f} did not download ({e}).")
            shutil.rmtree(tmp, ignore_errors=True)
            continue
        final = os.path.join(out, mid)
        shutil.rmtree(final, ignore_errors=True)
        shutil.move(os.path.join(tmp, mid), final)
        shutil.rmtree(tmp, ignore_errors=True)
        kept.append(m)
        print(f"  carried over {m.get('label', mid)} ({len(files)} files)")
    with open(os.path.join(out, "manifest.json"), "w", encoding="utf-8") as fh:
        json.dump({"months": kept, "latest": kept[-1]["id"] if kept else None}, fh)
    print(f"Carried over {len(kept)} of {len(months)} published month(s).")


if __name__ == "__main__":
    main()
