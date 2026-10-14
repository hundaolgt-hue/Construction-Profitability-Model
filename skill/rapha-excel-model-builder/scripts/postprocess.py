"""Inject cached values (computed by LibreOffice + brute-force data tables) into the
openpyxl-built workbook, so the file shows numbers in any viewer while keeping every
formula, Data Table, chart and style exactly as built. Excel recalculates on open."""
import re
import zipfile
from xml.sax.saxutils import escape
from openpyxl.utils import range_boundaries, get_column_letter as CL


def _sheet_files(z):
    wbx = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    rmap.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', rels)})
    out = {}
    for name, rid in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbx):
        tgt = rmap[rid]
        tgt = tgt.lstrip("/")
        out[name.replace("&amp;", "&")] = tgt if tgt.startswith("xl/") else "xl/" + tgt
    return out


def _vxml(v):
    if v is None:
        return None, None
    if isinstance(v, bool):
        return "b", "1" if v else "0"
    if isinstance(v, (int, float)):
        return None, repr(float(v)) if isinstance(v, float) else str(v)
    if hasattr(v, "toordinal"):  # datetime -> serial
        import datetime as dt
        base = dt.datetime(1899, 12, 30)
        d = v if isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day)
        return None, repr((d - base).total_seconds() / 86400)
    s = str(v)
    if s in ("#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NULL!", "#NUM!", "#N/A"):
        return "e", s
    return "str", escape(s)


def inject(src, dst, values, dt_ranges=()):
    """values: {sheet: {coord: value}}"""
    zin = zipfile.ZipFile(src)
    files = _sheet_files(zin)
    inv = {v: k for k, v in files.items()}
    dt_cells = {}
    for sh, ref in dt_ranges:
        c1, r1, c2, r2 = range_boundaries(ref)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                dt_cells.setdefault(sh, set()).add(f"{CL(c)}{r}")
    stats = {"formula": 0, "dt": 0}
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename in inv:
            sh = inv[item.filename]
            sv = values.get(sh, {})
            xml = data.decode("utf-8")

            def rep_formula(m):
                ref, attrs, ftag = m.group(1), m.group(2) or "", m.group(3)
                t, v = _vxml(sv.get(ref))
                attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
                if v is None:
                    return f'<c r="{ref}"{attrs}>{ftag}</c>'
                stats["formula"] += 1
                ta = f' t="{t}"' if t else ""
                return f'<c r="{ref}"{attrs}{ta}>{ftag}<v>{v}</v></c>'

            xml = re.sub(r'<c r="([A-Z]+\d+)"((?:\s+[a-z]+="[^"]*")*)>(<f[^>]*/>|<f(?:\s[^>]*)?>[^<]*</f>)<v\s*/>\s*</c>',
                         rep_formula, xml)
            dset = dt_cells.get(sh, set())
            if dset:
                def rep_empty(m):
                    ref, attrs = m.group(1), m.group(2) or ""
                    if ref not in dset:
                        return m.group(0)
                    t, v = _vxml(sv.get(ref))
                    if v is None:
                        return m.group(0)
                    stats["dt"] += 1
                    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
                    ta = f' t="{t}"' if t else ""
                    return f'<c r="{ref}"{attrs}{ta}><v>{v}</v></c>'
                xml = re.sub(r'<c r="([A-Z]+\d+)"((?:\s+[a-z]+="[^"]*")*)\s*/>', rep_empty, xml)
            data = xml.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    return stats
