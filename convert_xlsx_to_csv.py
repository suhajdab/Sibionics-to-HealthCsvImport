#!/usr/bin/env python3
import sys
import zipfile
import xml.etree.ElementTree as ET
import csv
import re
import os

def reformat_timestamp(s):
    s = s.strip()
    if not s:
        return s
    parts = s.split()
    # date part
    date_part = parts[0]
    time_part = parts[1] if len(parts) > 1 else ''
    ampm = None
    tz_part = None
    for p in parts[2:]:
        up = p.upper()
        if up in ('AM', 'PM') and ampm is None:
            ampm = up
        elif p.upper().startswith('GMT') or re.match(r'^[+-]\\d', p):
            tz_part = p
        else:
            if tz_part is None:
                tz_part = p
    # parse date MM-DD-YYYY
    try:
        m_str, d_str, y_str = date_part.split('-')
        month = int(m_str); day = int(d_str); year = int(y_str)
    except ValueError:
        return s
    # parse time hh:mm[:ss]
    hour = 0; minute = 0; sec = 0
    if ':' in time_part:
        tp = time_part.split(':')
        hour = int(tp[0]); minute = int(tp[1])
        if len(tp) > 2:
            sec = int(tp[2])
    else:
        try:
            hour = int(time_part)
        except:
            return s
    # adjust AM/PM
    if ampm == 'AM' and hour == 12:
        hour = 0
    elif ampm == 'PM' and hour < 12:
        hour += 12
    # parse timezone offset
    offset_str = ''
    if tz_part:
        offs = tz_part.upper().lstrip('GMT')
        if ':' in offs:
            sign = offs[0]
            hhmm = offs[1:].split(':')
            try:
                oh = int(hhmm[0]); om = int(hhmm[1])
                offset_str = f"{sign}{oh:02d}:{om:02d}"
            except:
                offset_str = ''
        else:
            try:
                oi = int(offs)
                sign = '+' if oi >= 0 else '-'
                offset_str = f"{sign}{abs(oi):02d}:00"
            except:
                offset_str = ''
    # build ISO timestamp
    dt = f"{year:04d}-{month:02d}-{day:02d} {hour:02d}:{minute:02d}:{sec:02d}"
    if offset_str:
        dt += f" {offset_str}"
    return dt

def col_to_index(col):
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx

def get_shared_strings(z):
    try:
        data = z.read('xl/sharedStrings.xml')
    except KeyError:
        return []
    sst = ET.fromstring(data)
    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    strings = []
    for si in sst.findall('ns:si', ns):
        texts = si.findall('.//ns:t', ns)
        text = ''.join([t.text or '' for t in texts])
        strings.append(text)
    return strings

def get_sheets(z):
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    nsr = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
    rel_map = {rel.get('Id'): rel.get('Target')
               for rel in rels.findall('r:Relationship', nsr)}
    sheets = []
    for sheet in wb.findall('ns:sheets/ns:sheet', ns):
        name = sheet.get('name')
        rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        target = rel_map.get(rid)
        if target:
            sheets.append((name, os.path.join('xl', target)))
    return sheets

def parse_sheet(z, path, shared_strings):
    data = z.read(path)
    tree = ET.fromstring(data)
    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    rows_data = []
    for row in tree.findall('ns:sheetData/ns:row', ns):
        row_data = {}
        max_idx = 0
        for c in row.findall('ns:c', ns):
            r = c.get('r')
            col = ''.join(filter(str.isalpha, r))
            idx = col_to_index(col)
            if idx > max_idx:
                max_idx = idx
            v = c.find('ns:v', ns)
            if v is None:
                value = ''
            else:
                value = v.text or ''
                if c.get('t') == 's':
                    try:
                        value = shared_strings[int(value)]
                    except Exception:
                        pass
            row_data[idx] = value
        rows = [row_data.get(i, '') for i in range(1, max_idx+1)]
        rows_data.append(rows)
    return rows_data

def main():
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <input.xlsx> [output_dir]")
        sys.exit(1)
    xlsx_path = sys.argv[1]
    out_dir = sys.argv[2] if len(sys.argv) > 2 else '.'
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        shared_strings = get_shared_strings(z)
        sheets = get_sheets(z)
        if not sheets:
            print("No sheets found.")
            sys.exit(1)
        for name, path in sheets:
            rows = parse_sheet(z, path, shared_strings)
            safe_name = ''.join(c if c.isalnum() else '_' for c in name)
            # Split sensor glucose records into per-day files
            if rows and len(rows[0]) >= 2 and 'sensor' in rows[0][1].lower() and 'reading' in rows[0][1].lower():
                header = rows[0]
                m = re.search(r"\(([^)]+)\)", header[1])
                unit = m.group(1) if m else ''
                grouped = {}
                for row in rows[1:]:
                    if not row or (len(row) >= 2 and row[0] == '' and row[1] == ''):
                        continue
                    ts = reformat_timestamp(row[0])
                    val = row[1] if len(row) > 1 else ''
                    date = ts[:10]
                    grouped.setdefault(date, []).append((ts, val))
                for date in sorted(grouped):
                    file_name = f"{safe_name}_{date}.csv"
                    out_path = os.path.join(out_dir, file_name)
                    with open(out_path, 'w', newline='') as f:
                        writer = csv.writer(f)
                        writer.writerow(['Timestamp', f'Blood Glucose ({unit})' if unit else 'Blood Glucose'])
                        for ts, val in grouped[date]:
                            writer.writerow([ts, val])
                    print(f"Wrote {out_path}")
            else:
                out_path = os.path.join(out_dir, f"{safe_name}.csv")
                with open(out_path, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerows(rows)
                print(f"Wrote {out_path}")

if __name__ == '__main__':
    main()