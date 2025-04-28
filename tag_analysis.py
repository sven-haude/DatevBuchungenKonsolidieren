#!/usr/bin/env python3
"""
tag_analysis.py  –  Wandelt einen Bank-RTF/CSV-Export in eine konsolidierte
Excel-Tabelle mit Monats-/Gruppensummen (inkl. D_, S_, G_-Gesamten).

Aufruf   :  python tag_analysis.py  <eingabeDatei>  [--out <excelDatei>]
Beispiel :  python tag_analysis.py  test1.rtf
"""

import argparse, re, sys, pathlib, datetime, pandas as pd, numpy as np

# ---------- Parsing-Funktionen ----------
def parse_german(num: str) -> float:
    if pd.isna(num): return np.nan
    s = str(num).strip()
    if not s: return np.nan
    sign = -1 if s.startswith('-') else 1
    s = s.lstrip('+-')
    if s.endswith('-'):
        sign = -1
        s = s[:-1]
    s = s.replace('.', '').replace(',', '.')
    return sign * float(s)

def parse_dot(num: str) -> float:
    if pd.isna(num): return np.nan
    s = str(num).strip()
    if not s: return np.nan
    sign = -1 if s.startswith('-') else 1
    s = s.lstrip('+-')
    if s.endswith('-'):
        sign = -1
        s = s[:-1]
    s = s.replace(',', '')
    return sign * float(s)

# ---------- RTF-Loader ----------
def load_rtf_semicol(path: pathlib.Path, min_cols: int = 6) -> pd.DataFrame:
    headers, rows, reading = None, [], False
    with path.open('r', encoding='utf-8-sig', errors='ignore') as f:
        for line in f:
            txt = line.rstrip('\r\n').rstrip('\\')
            if not reading:
                if ';' in txt:
                    parts = txt.split(';')
                    if len(parts) >= min_cols:
                        headers, reading = parts, True
            else:
                if txt.strip().startswith('}'):          # Ende der RTF-Tabelle
                    break
                if not txt.strip():
                    continue
                parts = txt.split(';')
                # auf Header-Breite normalisieren
                if len(parts) < len(headers):
                    parts += [''] * (len(headers) - len(parts))
                elif len(parts) > len(headers):
                    parts = parts[:len(headers)]
                rows.append(parts)
    if headers is None:
        raise ValueError("Keine Semikolon-Tabelle in der Datei gefunden.")
    return pd.DataFrame(rows, columns=headers)

# ---------- Hauptlogik ----------
def analyse_file(infile: pathlib.Path, outfile: pathlib.Path) -> None:
    # CSV direkt laden, sonst RTF auspacken
    if infile.suffix.lower() == '.csv':
        df = pd.read_csv(infile, sep=';', engine='python')
    else:
        df = load_rtf_semicol(infile)

    if 'Buchungsdatum' not in df.columns or 'Betrag in EUR' not in df.columns:
        raise KeyError("Spalten 'Buchungsdatum' oder 'Betrag in EUR' fehlen.")

    df['row_amount']     = df['Betrag in EUR'].apply(parse_german)
    df['Datum_dt']       = pd.to_datetime(df['Buchungsdatum'],
                                          format='%d.%m.%Y', errors='coerce')
    df.dropna(subset=['Datum_dt'], inplace=True)
    df['Monat']          = df['Datum_dt'].dt.to_period('M').astype(str)

    tag_re = re.compile(r'\['
                        r'(?P<base>[A-Za-z])'
                        r'(?:\s*:\s*(?P<num>[0-9\.,\-\+]+))?'
                        r'(?:\s*,\s*(?P<label>[^\]]+?))?'
                        r'\]')

    records = []
    for _, row in df.iterrows():
        note = str(row.get('Notiz', ''))
        for m in tag_re.finditer(note):
            base, num, label = m.group('base'), m.group('num'), m.group('label')
            grp = base if label is None else f"{base}, {label.strip()}"
            row_amt = row['row_amount']
            if pd.isna(row_amt): continue
            row_sign = -1 if row_amt < 0 else  1
            if num:
                explicit = num.lstrip().startswith(('+', '-')) \
                           or num.rstrip().endswith(('+', '-'))
                amt = parse_dot(num) if explicit else row_sign * abs(parse_dot(num))
            else:
                amt = row_amt
            records.append({'Monat': row['Monat'], 'Gruppe': grp, 'Betrag': amt})

    if not records:
        raise RuntimeError("Keine Tags gefunden – Abbruch.")

    detail = pd.DataFrame(records)
    pivot  = (detail
              .pivot_table(index='Monat', columns='Gruppe',
                           values='Betrag', aggfunc='sum', fill_value=0.0)
              .sort_index())

    # Sammelspalten D/S/G
    for prefix in ('D', 'S', 'G'):
        cols = [c for c in pivot.columns if c.startswith(prefix)]
        if cols:
            pivot[f'{prefix}_Summe'] = pivot[cols].sum(axis=1)

    # Sortierung
    agg_cols   = [c for c in pivot.columns if c.endswith('_Summe')]
    other_cols = sorted([c for c in pivot.columns if c not in agg_cols],
                        key=lambda x: (x[0], x.lower()))
    pivot      = pivot[other_cols + agg_cols].round(2)
    pivot.loc['Gesamt'] = pivot.sum(numeric_only=True, axis=0)
    pivot = pivot.reindex(['Gesamt', *pivot.index[:-1]])  # Gesamt nach oben ziehen

    outfile.parent.mkdir(parents=True, exist_ok=True)
    pivot.reset_index().to_excel(outfile, index=False)
    print(f"✔️  Excel geschrieben: {outfile}")

# ---------- CLI ----------
def main():
    ap = argparse.ArgumentParser(description="Bank-RTF/CSV → Monats-Excel")
    ap.add_argument('input', help="RTF- oder CSV-Datei mit ;-getrennter Tabelle")
    ap.add_argument('--out', '-o', help="Ausgabe-Excel",
                    default=None)
    args = ap.parse_args()

    src = pathlib.Path(args.input).expanduser().resolve()
    if not src.exists():
        print(f"Datei nicht gefunden: {src}", file=sys.stderr); sys.exit(1)

    if args.out:
        out = pathlib.Path(args.out).expanduser().resolve()
    else:
        stem = src.stem.replace('.', '_')
        out  = src.with_name(f'{stem}_MonatsSummen.xlsx')

    analyse_file(src, out)

if __name__ == '__main__':
    main()
