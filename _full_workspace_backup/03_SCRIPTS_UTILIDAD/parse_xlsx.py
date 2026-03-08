import glob
import zipfile
import xml.etree.ElementTree as ET
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='ignore')
except Exception:
    pass

def col_to_idx(col):
    n = 0
    for ch in col:
        if 'A' <= ch <= 'Z':
            n = n*26 + (ord(ch)-64)
    return n

xlsx = glob.glob(r'C:\Users\elrub\Downloads\MODELO 01*SUBVENCIONES.xlsx')
if not xlsx:
    raise SystemExit('No file found')
path = xlsx[0]
print('FILE found')

with zipfile.ZipFile(path, 'r') as z:
    wb_xml = ET.fromstring(z.read('xl/workbook.xml'))
    ns = {'x':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

    rels_xml = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rels = {}
    for rel in rels_xml:
        rid = rel.attrib.get('Id')
        target = rel.attrib.get('Target')
        if rid and target:
            rels[rid] = target

    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        sst = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in sst.findall('.//x:si', ns):
            txt = ''.join(t.text or '' for t in si.findall('.//x:t', ns))
            shared.append(txt)

    sheets = wb_xml.find('x:sheets', ns)
    for s in sheets.findall('x:sheet', ns):
        name = s.attrib.get('name')
        rid = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        target = rels.get(rid, '')
        if not target.startswith('xl/'):
            target = 'xl/' + target.lstrip('/')

        ws = ET.fromstring(z.read(target))
        rows = ws.findall('.//x:sheetData/x:row', ns)
        max_col = 0
        for c in ws.findall('.//x:sheetData/x:row/x:c', ns):
            ref = c.attrib.get('r','A1')
            col = ''.join(ch for ch in ref if ch.isalpha())
            max_col = max(max_col, col_to_idx(col))

        headers = {}
        if rows:
            first = rows[0]
            for c in first.findall('x:c', ns):
                ref = c.attrib.get('r','A1')
                col = ''.join(ch for ch in ref if ch.isalpha())
                idx = col_to_idx(col)
                t = c.attrib.get('t')
                v_el = c.find('x:v', ns)
                val = ''
                if v_el is not None and v_el.text is not None:
                    raw = v_el.text
                    if t == 's':
                        try:
                            val = shared[int(raw)]
                        except Exception:
                            val = raw
                    else:
                        val = raw
                else:
                    is_el = c.find('x:is', ns)
                    if is_el is not None:
                        val = ''.join(tn.text or '' for tn in is_el.findall('.//x:t', ns))
                headers[idx] = val

        ordered = [headers.get(i, '') for i in range(1, min(max_col, 25)+1)]
        print('---', name, 'rows', len(rows), 'maxCol', max_col)
        print(ordered)
