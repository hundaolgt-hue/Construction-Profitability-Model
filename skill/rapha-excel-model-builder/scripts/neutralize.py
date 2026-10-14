"""Make buyer copies: neutral author/title metadata on .xlsx and .pdf files (no internal names, no tool fingerprints).

usage: python neutralize.py OUT_DIR file1.xlsx [file2.pdf ...] [--author Rapha] [--title-prefix "Rapha ..."]
Each output keeps the same file name inside OUT_DIR. Titles are derived from the file name unless --title is given per file.
"""
import argparse
import os
import re
import zipfile


def xlsx(src, dst, title, author):
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties '
            'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{title}</dc:title><dc:creator>{author}</dc:creator><cp:lastModifiedBy>{author}</cp:lastModifiedBy>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>')
    zin = zipfile.ZipFile(src)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        d = zin.read(it.filename)
        if it.filename == "docProps/core.xml":
            d = core.encode()
        elif it.filename == "docProps/app.xml":
            d = re.sub(rb"<Application>[^<]*</Application>", b"<Application>Microsoft Excel</Application>", d)
        zout.writestr(it, d)
    zout.close()


def pdf(src, dst, title, author):
    from pypdf import PdfReader, PdfWriter
    r = PdfReader(src)
    w = PdfWriter()
    for p in r.pages:
        w.add_page(p)
    w.add_metadata({"/Title": title, "/Author": author, "/Creator": author, "/Producer": author})
    w.write(dst)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("out_dir")
    ap.add_argument("files", nargs="+")
    ap.add_argument("--author", default="Rapha")
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    for f in a.files:
        name = os.path.basename(f)
        title = os.path.splitext(name)[0].replace("_", " ")
        dst = os.path.join(a.out_dir, name)
        (xlsx if f.lower().endswith(".xlsx") else pdf)(f, dst, title, a.author)
        print("wrote", dst)
