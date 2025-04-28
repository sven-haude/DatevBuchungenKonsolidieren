"""
Streamlit-Webanwendung: Konsolidierte Monats- und Gruppensummen aus Bank-Exporten
===============================================================================

Der Benutzer kann entweder eine RTF/CSV-Datei hochladen **oder** den
Tabellen-Inhalt (Semikolon-getrennt) in ein Textfeld kopieren. Die App parst
die Daten, erkennt die in eckigen Klammern gesetzten Tags (z. B. `[D, Miete]`)
und erstellt eine Pivot-Tabelle mit Monats- und Gruppensummen sowie den
Aggregat-Spalten `D_Summe`, `S_Summe`, `G_Summe`. Anschließend wird eine
Excel-Datei erzeugt, die per Download-Button heruntergeladen werden kann.
"""

import io
import re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

# ----------------------------------------------------------- Parsing-Funktionen

def parse_german(num: str) -> float:
    """Wandelt eine deutsche Zahl (1.234,56 / -1.234,56-) in float."""
    if pd.isna(num):
        return np.nan
    s = str(num).strip()
    if not s:
        return np.nan
    sign = -1 if s.startswith("-") else 1
    s = s.lstrip("+-")
    if s.endswith("-"):
        sign = -1
        s = s[:-1]
    s = s.replace(".", "").replace(",", ".")
    return sign * float(s)


def parse_dot(num: str) -> float:
    """Wandelt 1,234.56 / -1,234.56- (engl. Format) in float."""
    if pd.isna(num):
        return np.nan
    s = str(num).strip()
    if not s:
        return np.nan
    sign = -1 if s.startswith("-") else 1
    s = s.lstrip("+-")
    if s.endswith("-"):
        sign = -1
        s = s[:-1]
    s = s.replace(",", "")
    return sign * float(s)


# ----------------------------------------------------------- RTF-Loader

def _load_rtf_semicol(lines: list[str], min_cols: int = 6) -> pd.DataFrame:
    """Extrak­tiert die erste Semikolon-Tabelle aus RTF-Textzeilen."""
    headers, rows, reading = None, [], False
    for line in lines:
        txt = line.rstrip("\r\n").rstrip("\\")
        if not reading:
            if ";" in txt:
                parts = txt.split(";")
                if len(parts) >= min_cols:
                    headers, reading = parts, True
        else:
            if txt.strip().startswith("}"):  # Ende der RTF-Tabelle
                break
            if not txt.strip():
                continue
            parts = txt.split(";")
            if len(parts) < len(headers):
                parts += [""] * (len(headers) - len(parts))
            elif len(parts) > len(headers):
                parts = parts[: len(headers)]
            rows.append(parts)
    if headers is None:
        raise ValueError("Keine Semikolon-Tabelle in der Datei gefunden.")
    return pd.DataFrame(rows, columns=headers)


# ----------------------------------------------------------- Tag-Analyse-Logik

tag_re = re.compile(
    r"\["
    r"(?P<base>[A-Za-z])"
    r"(?:\s*:\s*(?P<num>[0-9\.,\-\+]+))?"
    r"(?:\s*,\s*(?P<label>[^\]]+?))?"
    r"\]",
)


def analyse(df_raw: pd.DataFrame) -> pd.DataFrame:
    """Erstellt Pivot mit Monats- und Gruppensummen inkl. Sammel-Spalten."""
    if "Buchungsdatum" not in df_raw.columns or "Betrag in EUR" not in df_raw.columns:
        raise KeyError("Spalten 'Buchungsdatum' oder 'Betrag in EUR' fehlen.")

    df = df_raw.copy()
    df["row_amount"] = df["Betrag in EUR"].apply(parse_german)
    df["Datum_dt"] = pd.to_datetime(df["Buchungsdatum"], format="%d.%m.%Y", errors="coerce")
    df.dropna(subset=["Datum_dt"], inplace=True)
    df["Monat"] = df["Datum_dt"].dt.to_period("M").astype(str)

    records: list[dict[str, object]] = []
    for _, row in df.iterrows():
        note = str(row.get("Notiz", ""))
        for m in tag_re.finditer(note):
            base, num, label = m.group("base"), m.group("num"), m.group("label")
            grp = base if label is None else f"{base}, {label.strip()}"
            row_amt = row["row_amount"]
            if pd.isna(row_amt):
                continue
            row_sign = -1 if row_amt < 0 else 1
            if num:
                explicit = num.lstrip().startswith(("+", "-")) or num.rstrip().endswith(("+", "-"))
                amt = parse_dot(num) if explicit else row_sign * abs(parse_dot(num))
            else:
                amt = row_amt
            records.append({"Monat": row["Monat"], "Gruppe": grp, "Betrag": amt})

    if not records:
        raise RuntimeError("Keine Tags gefunden – Abbruch.")

    detail = pd.DataFrame(records)
    pivot = (
        detail.pivot_table(index="Monat", columns="Gruppe", values="Betrag", aggfunc="sum", fill_value=0.0)
        .sort_index()
    )

    # Sammelspalten D/S/G
    for prefix in ("D", "S", "G"):
        cols = [c for c in pivot.columns if c.startswith(prefix)]
        if cols:
            pivot[f"{prefix}_Summe"] = pivot[cols].sum(axis=1)

    # Sortierung
    agg_cols = [c for c in pivot.columns if c.endswith("_Summe")]
    other_cols = sorted([c for c in pivot.columns if c not in agg_cols], key=lambda x: (x[0], x.lower()))
    pivot = pivot[other_cols + agg_cols].round(2)

    pivot.loc["Gesamt"] = pivot.sum(numeric_only=True, axis=0)
    pivot = pivot.reindex(["Gesamt", *pivot.index[:-1]])  # Gesamt nach oben

    return pivot


# ----------------------------------------------------------- UI-Funktionen

def read_input(uploaded_file, pasted_text: str) -> pd.DataFrame | None:
    """Liest Datei-Upload oder Text-Eingabe ein und gibt DataFrame zurück."""
    if uploaded_file is not None:
        name = uploaded_file.name.lower()
        data = uploaded_file.read()
        if name.endswith(".csv"):
            return pd.read_csv(io.BytesIO(data), sep=";", engine="python")
        elif name.endswith(".rtf"):
            lines = io.BytesIO(data).read().decode("utf-8-sig", "ignore").splitlines()
            return _load_rtf_semicol(lines)
        else:
            st.error("Nur *.rtf oder *.csv werden unterstützt.")
            return None
    elif pasted_text.strip():
        # Versuchen, RTF-Header zu erkennen, sonst als CSV/Semikolon-Tabelle interpretieren
        if pasted_text.lstrip().startswith("{\\rtf"):
            return _load_rtf_semicol(pasted_text.splitlines())
        else:
            return pd.read_csv(io.StringIO(pasted_text), sep=";", engine="python")
    return None


# ----------------------------------------------------------- Haupt-Anwendung

def main() -> None:
    st.set_page_config(page_title="Tag-Analyse", layout="centered")
    st.title("🔖 Bank-Tag-Analyse → Monats-Excel")

    st.markdown(
        """Lade einen **RTF- oder CSV-Export** deiner Bank herunter und tagge die
        Zeilen z. B. mit `[D, Miete]`, `[S, Gehalt]` usw. Anschließend kannst du
        hier die Datei hochladen **oder** den Tabellen-Inhalt in das Textfeld
        kopieren. Die App erstellt eine Excel-Datei mit Monats- und
        Gruppensummen, inklusive Sammelspalten **D/S/G**.""",
    )

    uploaded_file = st.file_uploader("Datei hochladen (.rtf oder .csv)")
    st.markdown("**oder**")
    pasted_text = st.text_area("Tabellen-Text einkopieren (optional)", height=200)

    if st.button("Analysieren und Excel erzeugen"):
        try:
            df_source = read_input(uploaded_file, pasted_text)
            if df_source is None:
                st.warning("Bitte eine Datei hochladen oder Text einkopieren.")
                st.stop()

            pivot = analyse(df_source)
            st.success("Analyse erfolgreich!")
            st.dataframe(pivot, use_container_width=True)

            # Excel in-memory erstellen
            buf = io.BytesIO()
            pivot.reset_index().to_excel(buf, index=False)
            buf.seek(0)
            st.download_button(
                label="📥 Excel herunterladen",
                data=buf,
                file_name="MonatsSummen.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.exception(e)


if __name__ == "__main__":
    main()
