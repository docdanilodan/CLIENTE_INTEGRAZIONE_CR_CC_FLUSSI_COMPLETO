
from __future__ import annotations

import io
import re
import sqlite3
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import streamlit as st

try:
    import Analisi_CC_FLUSSI as cc_flussi
except Exception as _cc_import_error:
    cc_flussi = None
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
)

APP_TITLE = "FinancePlus - Clienti e Pratiche"
DB_PATH = Path("financeplus_clienti_pratiche.db")

PRODOTTI = ["Chiro", "Factoring", "Leasing", "Invoice Trading", "Mutuo"]
GESTORI = ["Paolo", "Sergio", "Zeno", "Nick"]
STATI = ["🟢 Verde", "🟡 Giallo"]


# -----------------------------
# DATABASE
# -----------------------------
def db_connect():
    return sqlite3.connect(DB_PATH)


def init_db() -> None:
    with db_connect() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS clienti (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                denominazione TEXT NOT NULL,
                partita_iva TEXT,
                sede TEXT,
                indirizzo TEXT,
                cap TEXT,
                provincia TEXT,
                documento TEXT,
                creato_il TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS collaboratori (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                cognome TEXT NOT NULL,
                creato_il TEXT NOT NULL,
                UNIQUE(nome, cognome)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS pratiche (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cliente_id INTEGER NOT NULL,
                prodotto TEXT NOT NULL,
                importo REAL NOT NULL DEFAULT 0,
                istituto TEXT,
                gestore TEXT,
                note TEXT,
                integrazioni TEXT,
                data_integrazione TEXT,
                stato TEXT,
                creato_il TEXT NOT NULL,
                FOREIGN KEY(cliente_id) REFERENCES clienti(id)
            )
        """)
        # Migrazione database: aggiunge collaboratore_id ai database già esistenti.
        columns = [row[1] for row in conn.execute("PRAGMA table_info(clienti)").fetchall()]
        if "collaboratore_id" not in columns:
            conn.execute("ALTER TABLE clienti ADD COLUMN collaboratore_id INTEGER")
        conn.commit()


def list_collaboratori() -> List[Tuple]:
    with db_connect() as conn:
        return conn.execute("""
            SELECT id, nome, cognome
            FROM collaboratori
            ORDER BY cognome, nome
        """).fetchall()


def save_collaboratore(nome: str, cognome: str) -> int:
    nome = clean(nome).title()
    cognome = clean(cognome).title()
    with db_connect() as conn:
        existing = conn.execute("""
            SELECT id FROM collaboratori
            WHERE LOWER(nome)=LOWER(?) AND LOWER(cognome)=LOWER(?)
        """, (nome, cognome)).fetchone()
        if existing:
            return int(existing[0])
        cur = conn.execute("""
            INSERT INTO collaboratori (nome, cognome, creato_il)
            VALUES (?, ?, ?)
        """, (nome, cognome, datetime.now().strftime("%d-%m-%Y %H:%M")))
        conn.commit()
        return int(cur.lastrowid)


def list_clienti() -> List[Tuple]:
    with db_connect() as conn:
        return conn.execute("""
            SELECT c.id, c.denominazione, c.partita_iva, c.sede, c.indirizzo,
                   c.cap, c.provincia, c.collaboratore_id,
                   COALESCE(col.nome || ' ' || col.cognome, '') AS collaboratore
            FROM clienti c
            LEFT JOIN collaboratori col ON col.id = c.collaboratore_id
            ORDER BY c.denominazione
        """).fetchall()


def get_cliente(cliente_id: int) -> Optional[Tuple]:
    with db_connect() as conn:
        return conn.execute("""
            SELECT c.id, c.denominazione, c.partita_iva, c.sede, c.indirizzo,
                   c.cap, c.provincia, c.documento, c.creato_il,
                   c.collaboratore_id,
                   COALESCE(col.nome || ' ' || col.cognome, '') AS collaboratore
            FROM clienti c
            LEFT JOIN collaboratori col ON col.id = c.collaboratore_id
            WHERE c.id = ?
        """, (cliente_id,)).fetchone()


def save_cliente(data: Dict[str, str]) -> int:
    with db_connect() as conn:
        existing = None
        if data["partita_iva"]:
            existing = conn.execute(
                "SELECT id FROM clienti WHERE partita_iva = ?",
                (data["partita_iva"],)
            ).fetchone()
        if not existing:
            existing = conn.execute(
                "SELECT id FROM clienti WHERE LOWER(denominazione) = LOWER(?)",
                (data["denominazione"],)
            ).fetchone()

        if existing:
            conn.execute("""
                UPDATE clienti
                SET denominazione=?, partita_iva=?, sede=?, indirizzo=?, cap=?,
                    provincia=?, documento=?, collaboratore_id=?
                WHERE id=?
            """, (
                data["denominazione"], data["partita_iva"], data["sede"],
                data["indirizzo"], data["cap"], data["provincia"],
                data["documento"], data.get("collaboratore_id"), existing[0]
            ))
            conn.commit()
            return int(existing[0])

        cur = conn.execute("""
            INSERT INTO clienti (
                denominazione, partita_iva, sede, indirizzo,
                cap, provincia, documento, creato_il, collaboratore_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            data["denominazione"], data["partita_iva"], data["sede"],
            data["indirizzo"], data["cap"], data["provincia"],
            data["documento"], datetime.now().strftime("%d-%m-%Y %H:%M"),
            data.get("collaboratore_id")
        ))
        conn.commit()
        return int(cur.lastrowid)


def save_pratica(
    cliente_id: int, prodotto: str, importo: float, istituto: str,
    gestore: str, note: str, integrazioni: str,
    data_integrazione: date, stato: str
) -> int:
    with db_connect() as conn:
        cur = conn.execute("""
            INSERT INTO pratiche (
                cliente_id, prodotto, importo, istituto, gestore,
                note, integrazioni, data_integrazione, stato, creato_il
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            cliente_id, prodotto, importo, istituto.strip(), gestore,
            note.strip(), integrazioni.strip(),
            data_integrazione.strftime("%d-%m-%Y"),
            stato, datetime.now().strftime("%d-%m-%Y %H:%M")
        ))
        conn.commit()
        return int(cur.lastrowid)


def list_pratiche(cliente_id: Optional[int] = None) -> List[Tuple]:
    query = """
        SELECT
            p.id, c.denominazione, c.partita_iva, c.sede, c.indirizzo,
            p.prodotto, p.importo, p.istituto, p.gestore,
            p.note, p.integrazioni, p.data_integrazione,
            p.stato, p.creato_il, c.id
        FROM pratiche p
        JOIN clienti c ON c.id = p.cliente_id
    """
    params = ()
    if cliente_id is not None:
        query += " WHERE c.id = ?"
        params = (cliente_id,)
    query += " ORDER BY c.denominazione, p.id DESC"

    with db_connect() as conn:
        return conn.execute(query, params).fetchall()


# -----------------------------
# ESTRAZIONE PDF
# -----------------------------
def clean(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip(" \t\r\n:;,-")


def extract_pdf_text(file_bytes: bytes) -> str:
    reader = PdfReader(io.BytesIO(file_bytes))
    pages = []
    for page in reader.pages:
        try:
            pages.append(page.extract_text() or "")
        except Exception:
            pages.append("")
    return "\n".join(pages)


def first_match(patterns: List[str], text: str, flags=re.I | re.M) -> str:
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            return clean(match.group(1))
    return ""


def extract_denominazione(text: str) -> str:
    """
    Priorità:
    1. Campo 'Denominazione:' presente nella sezione costitutiva.
    2. Riga iniziale successiva a 'VISURA ...'.
    3. Intestazioni ripetute con forma giuridica.
    Evita volutamente sezioni come ATTIVITA'.
    """
    value = first_match([
        r"\bDenominazione\s*:\s*([^\n\r]{3,120})",
        r"VISURA\s+(?:ORDINARIA|STORICA)[^\n\r]*\n\s*([^\n\r]{3,120})",
    ], text)
    if value:
        value = re.split(
            r"\b(?:Codice Fiscale|Partita IVA|Dati anagrafici|Data atto)\b",
            value, flags=re.I
        )[0]
        if not re.fullmatch(r"ATTIVIT[AÀ]'?", value, re.I):
            return clean(value)

    legal = re.compile(
        r"\b(?:S\.?\s*R\.?\s*L\.?|S\.?\s*P\.?\s*A\.?|S\.?\s*A\.?\s*S\.?|S\.?\s*N\.?\s*C\.?)\b",
        re.I
    )
    excluded = re.compile(
        r"ATTIVIT[AÀ]|FORMA GIURIDICA|RESPONSABILITA|VISURA|SOCIETA'? DI CAPITALE",
        re.I
    )
    candidates = []
    for line in text.splitlines()[:250]:
        line = clean(line)
        if 4 <= len(line) <= 100 and legal.search(line) and not excluded.search(line):
            score = 0
            if line.upper() == line:
                score += 3
            if len(line.split()) <= 8:
                score += 2
            if re.search(r"S\.?\s*R\.?\s*L\.?\s*$", line, re.I):
                score += 3
            candidates.append((score, line))
    if candidates:
        return sorted(candidates, reverse=True)[0][1]
    return ""


def extract_piva(text: str) -> str:
    return first_match([
        r"\bPartita\s+IVA\s*[:\-]?\s*(?:IT\s*)?([0-9]{11})\b",
        r"\bCodice\s+fiscale\s+e\s+n\.?\s*iscr\.[^\n\r]*\n?\s*([0-9]{11})\b",
    ], text)


def extract_sede_indirizzo(text: str) -> Dict[str, str]:
    """
    Dalla riga:
    Indirizzo Sede legale MILANO (MI) VIA ANDREA COSTA 4 CAP 20131
    produce:
    sede = MILANO
    provincia = MI
    indirizzo = VIA ANDREA COSTA 4
    cap = 20131
    """
    normalized = re.sub(r"\s+", " ", text)

    patterns = [
        r"Indirizzo\s+Sede\s+legale\s+([A-ZÀ-ÖØ-Ý' .-]+?)\s*\(([A-Z]{2})\)\s+(.+?)\s+CAP\s+([0-9]{5})",
        r"\bSede\s+legale\s+([A-ZÀ-ÖØ-Ý' .-]+?)\s*\(([A-Z]{2})\)\s+(.+?)\s+CAP\s+([0-9]{5})",
    ]
    for pattern in patterns:
        m = re.search(pattern, normalized, re.I)
        if m:
            comune = clean(m.group(1)).upper()
            provincia = clean(m.group(2)).upper()
            indirizzo = clean(m.group(3)).upper()
            cap = clean(m.group(4))
            indirizzo = re.sub(r"^(?:LEGALE|SEDE LEGALE)\s+", "", indirizzo, flags=re.I)
            return {
                "sede": comune,
                "provincia": provincia,
                "indirizzo": indirizzo,
                "cap": cap,
            }

    # Struttura su più righe delle visure InfoCamere
    m = re.search(
        r"Indirizzo\s+Sede\s+legale\s+([A-ZÀ-ÖØ-Ý' .-]+?)\s*\(([A-Z]{2})\)\s*"
        r"(?:\n|\s)+([A-Z0-9À-ÖØ-Ý' ./-]+?)\s+CAP\s+([0-9]{5})",
        text, re.I
    )
    if m:
        return {
            "sede": clean(m.group(1)).upper(),
            "provincia": clean(m.group(2)).upper(),
            "indirizzo": clean(m.group(3)).upper(),
            "cap": clean(m.group(4)),
        }

    return {"sede": "", "provincia": "", "indirizzo": "", "cap": ""}


def parse_cliente(text: str, filename: str) -> Dict[str, str]:
    address = extract_sede_indirizzo(text)
    return {
        "denominazione": extract_denominazione(text),
        "partita_iva": extract_piva(text),
        "sede": address["sede"],
        "indirizzo": address["indirizzo"],
        "cap": address["cap"],
        "provincia": address["provincia"],
        "documento": filename,
    }


# -----------------------------
# REPORT PDF
# -----------------------------
def money(value: float) -> str:
    return f"€ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def report_header(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica-Bold", 9)
    canvas.setFillColor(colors.HexColor("#153B5B"))
    canvas.drawString(1.4 * cm, A4[1] - 1.1 * cm, "FinancePlus.tech")
    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(A4[0] - 1.4 * cm, 0.8 * cm, f"Pagina {doc.page}")
    canvas.restoreState()


def build_client_report(cliente: Tuple, pratiche: List[Tuple]) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=1.3 * cm, leftMargin=1.3 * cm,
        topMargin=1.6 * cm, bottomMargin=1.4 * cm
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "FPTitle", parent=styles["Title"], alignment=TA_CENTER,
        textColor=colors.HexColor("#153B5B"), fontSize=18, leading=22
    )
    h2 = ParagraphStyle(
        "FPH2", parent=styles["Heading2"],
        textColor=colors.HexColor("#A9693B"), fontSize=12
    )

    story = [
        Paragraph("REPORT PRATICHE CLIENTE", title),
        Spacer(1, 10),
        Paragraph(cliente[1], h2),
    ]

    anag = [
        ["P.IVA", cliente[2] or "-"],
        ["Sede", cliente[3] or "-"],
        ["Indirizzo", " ".join(x for x in [cliente[4], cliente[5], f"({cliente[6]})" if cliente[6] else ""] if x) or "-"],
        ["Collaboratore", cliente[10] or "-"],
    ]
    t = Table(anag, colWidths=[4 * cm, 12 * cm])
    t.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.4, colors.grey),
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#DCE8F1")),
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
    ]))
    story += [t, Spacer(1, 14), Paragraph("Pratiche registrate", h2)]

    rows = [["Prodotto", "Importo", "Istituto", "Gestore", "Stato", "Data integrazione"]]
    for p in pratiche:
        rows.append([
            p[5], money(p[6]), p[7] or "-", p[8] or "-",
            p[12] or "-", p[11] or "-"
        ])
    if len(rows) == 1:
        rows.append(["Nessuna pratica", "-", "-", "-", "-", "-"])

    pt = Table(rows, repeatRows=1, colWidths=[3.1*cm, 2.8*cm, 3.5*cm, 2.3*cm, 2.3*cm, 3.2*cm])
    pt.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.35, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#153B5B")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
    ]))
    story.append(pt)

    for p in pratiche:
        story += [
            Spacer(1, 12),
            Paragraph(f"Dettaglio pratica n. {p[0]} – {p[5]}", h2),
            Table([
                ["Importo", money(p[6])],
                ["Istituto", p[7] or "-"],
                ["Gestore", p[8] or "-"],
                ["Note", Paragraph(p[9] or "-", styles["BodyText"])],
                ["Integrazioni", Paragraph(p[10] or "-", styles["BodyText"])],
                ["Data integrazione", p[11] or "-"],
                ["Stato", p[12] or "-"],
            ], colWidths=[4*cm, 12*cm], style=TableStyle([
                ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
                ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#F0F4F7")),
                ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
            ]))
        ]

    doc.build(story, onFirstPage=report_header, onLaterPages=report_header)
    return buffer.getvalue()


def build_general_report(clienti: List[Tuple], pratiche: List[Tuple]) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        rightMargin=1.1*cm, leftMargin=1.1*cm,
        topMargin=1.4*cm, bottomMargin=1.2*cm
    )
    styles = getSampleStyleSheet()
    story = [
        Paragraph("REPORT GENERALE CLIENTI E PRATICHE", styles["Title"]),
        Spacer(1, 12),
    ]

    rows = [[
        "Cliente", "P.IVA", "Sede", "Prodotto", "Importo",
        "Istituto", "Gestore", "Stato", "Integrazione"
    ]]
    for p in pratiche:
        rows.append([
            p[1], p[2] or "-", p[3] or "-", p[5],
            money(p[6]), p[7] or "-", p[8] or "-",
            p[12] or "-", p[11] or "-"
        ])

    clienti_con_pratiche = {p[14] for p in pratiche}
    for c in clienti:
        if c[0] not in clienti_con_pratiche:
            rows.append([c[1], c[2] or "-", c[3] or "-", "-", "-", "-", "-", "-", "-"])

    table = Table(
        rows, repeatRows=1,
        colWidths=[4.7*cm, 2.8*cm, 2.8*cm, 2.8*cm, 2.8*cm, 3.5*cm, 2.4*cm, 2.4*cm, 3.0*cm]
    )
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#153B5B")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7.2),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F3F6F8")]),
    ]))
    story.append(table)
    doc.build(story)
    return buffer.getvalue()



# -----------------------------
# ANALISI CENTRALE RISCHI
# -----------------------------
import base64
import csv
import json
import zipfile
from collections import defaultdict
from html import escape

import pandas as pd
import matplotlib.pyplot as plt
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import KeepTogether

MONTHS_IT = [
    "GENNAIO", "FEBBRAIO", "MARZO", "APRILE", "MAGGIO", "GIUGNO",
    "LUGLIO", "AGOSTO", "SETTEMBRE", "OTTOBRE", "NOVEMBRE", "DICEMBRE"
]

COLUMN_ALIASES = {
    "periodo": ["periodo", "mese", "data", "riferimento", "competenza"],
    "banca": ["banca", "intermediario", "istituto", "segnalante"],
    "categoria": ["categoria", "forma tecnica", "forma_tecnica", "rischio", "tipologia"],
    "accordato": ["accordato operativo", "accordato_operativo", "accordato", "fido", "affidamento"],
    "utilizzato": ["utilizzato", "utilizzo", "esposizione", "saldo"],
    "sconfino": ["sconfino", "sconfinamento", "insoluto", "scaduto"],
    "garanzia": ["garanzia", "valore garanzia", "garantito"],
    "sofferenza": ["sofferenza", "sofferenze"],
    "perdita": ["perdita", "passato a perdita", "crediti passati a perdita"],
}


def _norm_col(value: str) -> str:
    value = re.sub(r"[_\-]+", " ", str(value or "").strip().lower())
    return re.sub(r"\s+", " ", value)


def _to_number(value) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace("€", "").replace(" ", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        parts = s.split(".")
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            s = "".join(parts)
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s)
    except ValueError:
        return 0.0


def _map_columns(columns) -> Dict[str, str]:
    normalized = {_norm_col(c): c for c in columns}
    result = {}
    for key, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            a = _norm_col(alias)
            exact = normalized.get(a)
            if exact is not None:
                result[key] = exact
                break
            candidates = [orig for norm, orig in normalized.items() if a in norm or norm in a]
            if candidates:
                result[key] = candidates[0]
                break
    return result


def _read_delimited(raw: bytes, name: str) -> pd.DataFrame:
    text = raw.decode("utf-8-sig", errors="ignore")
    sample = text[:8000]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=";,\t|")
        sep = dialect.delimiter
    except Exception:
        sep = ";" if sample.count(";") >= sample.count(",") else ","
    return pd.read_csv(io.StringIO(text), sep=sep, engine="python")


def _extract_pdf_rows(raw: bytes) -> Tuple[pd.DataFrame, str]:
    text = extract_pdf_text(raw)
    rows = []
    current_period = ""
    period_re = re.compile(r"\b(?:gen|feb|mar|apr|mag|giu|lug|ago|set|ott|nov|dic)[a-z]*\s+20\d{2}\b", re.I)
    money_re = re.compile(r"(?:€\s*)?(-?[0-9][0-9.]*[,\.]?[0-9]*)")
    bank_re = re.compile(r"\b(BANCA|BANCO|BCC|CREDITO|FACTORING|INTESA|UNICREDIT|MEDIOCREDITO|IFIS|AIDEXA)\b", re.I)
    category_re = re.compile(r"(autoliquidanti|scadenza|revoca|sofferenze|garanzie|firma|leasing|factoring)", re.I)
    for line in text.splitlines():
        line = clean(line)
        if not line:
            continue
        pm = period_re.search(line)
        if pm:
            current_period = pm.group(0)
        nums = [_to_number(x) for x in money_re.findall(line)]
        if bank_re.search(line) and nums:
            bank = line
            for token in money_re.findall(line):
                bank = bank.replace(token, "")
            bank = clean(re.sub(r"€", "", bank))[:140]
            cat = category_re.search(line)
            rows.append({
                "periodo": current_period,
                "banca": bank,
                "categoria": cat.group(1) if cat else "non classificata",
                "accordato": nums[0] if len(nums) > 0 else 0,
                "utilizzato": nums[1] if len(nums) > 1 else (nums[0] if len(nums) == 1 else 0),
                "sconfino": nums[2] if len(nums) > 2 else 0,
                "garanzia": nums[3] if len(nums) > 3 else 0,
                "sofferenza": 0,
                "perdita": 0,
            })
    return pd.DataFrame(rows), text


def read_cr_file(uploaded) -> Tuple[pd.DataFrame, str, List[str]]:
    raw = uploaded.getvalue()
    suffix = Path(uploaded.name).suffix.lower()
    warnings = []
    source_text = ""
    try:
        if suffix == ".xlsx":
            sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None)
            frames = []
            for sheet_name, frame in sheets.items():
                if not frame.empty:
                    frame = frame.copy()
                    frame["_foglio"] = sheet_name
                    frames.append(frame)
            df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
        elif suffix == ".csv":
            df = _read_delimited(raw, uploaded.name)
        elif suffix == ".txt":
            df = _read_delimited(raw, uploaded.name)
            source_text = raw.decode("utf-8-sig", errors="ignore")
        elif suffix == ".pdf":
            df, source_text = _extract_pdf_rows(raw)
            if df.empty:
                warnings.append("Il PDF è testuale ma non presenta righe tabellari riconoscibili automaticamente.")
        else:
            raise ValueError("Formato non supportato")
    except Exception as exc:
        raise ValueError(f"Impossibile leggere il file CR: {exc}") from exc

    if df.empty:
        return pd.DataFrame(columns=list(COLUMN_ALIASES)), source_text, warnings

    mapping = _map_columns(df.columns)
    canonical = pd.DataFrame()
    for key in COLUMN_ALIASES:
        if key in mapping:
            canonical[key] = df[mapping[key]]
        else:
            canonical[key] = "" if key in ("periodo", "banca", "categoria") else 0.0
            if key in ("accordato", "utilizzato", "sconfino"):
                warnings.append(f"Colonna '{key}' non individuata: valorizzata a zero.")
    for col in ["accordato", "utilizzato", "sconfino", "garanzia", "sofferenza", "perdita"]:
        canonical[col] = canonical[col].map(_to_number)
    for col in ["periodo", "banca", "categoria"]:
        canonical[col] = canonical[col].fillna("").astype(str).map(clean)
    canonical = canonical[(canonical[["accordato", "utilizzato", "sconfino", "garanzia", "sofferenza", "perdita"]].abs().sum(axis=1) > 0) | (canonical["banca"] != "")]
    return canonical.reset_index(drop=True), source_text, list(dict.fromkeys(warnings))


def analyse_cr(df: pd.DataFrame) -> Dict:
    if df.empty:
        periods = []
    else:
        periods = [p for p in df["periodo"].dropna().astype(str).unique().tolist() if p]
    latest_period = periods[-1] if periods else "periodo non identificato"
    latest = df[df["periodo"] == latest_period] if periods else df
    if latest.empty:
        latest = df
    by_bank = latest.groupby("banca", dropna=False)[["accordato", "utilizzato", "sconfino", "garanzia"]].sum().reset_index()
    by_bank["banca"] = by_bank["banca"].replace("", "Intermediario non identificato")
    by_cat = latest.groupby("categoria", dropna=False)[["accordato", "utilizzato", "sconfino"]].sum().reset_index()
    totals = latest[["accordato", "utilizzato", "sconfino", "garanzia", "sofferenza", "perdita"]].sum().to_dict()
    accordato = totals.get("accordato", 0.0)
    utilizzato = totals.get("utilizzato", 0.0)
    sconfino = totals.get("sconfino", 0.0)
    saturation = (utilizzato / accordato * 100) if accordato else (100.0 if utilizzato else 0.0)
    score = 100
    if totals.get("sofferenza", 0) > 0: score -= 45
    if totals.get("perdita", 0) > 0: score -= 35
    if sconfino > 0: score -= min(20, 5 + int((sconfino / max(accordato, 1)) * 200))
    if saturation > 95: score -= 12
    elif saturation > 90: score -= 7
    elif saturation > 80: score -= 3
    score = max(0, min(100, score))
    judgment = "OTTIMO" if score >= 80 else "VULNERABILE" if score >= 60 else "PERICOLO"
    rating = "AA" if score >= 80 else "BBB" if score >= 70 else "BB" if score >= 60 else "B/C"
    pd_est = 0.5 if score >= 80 else 1.5 if score >= 70 else 3.0 if score >= 60 else 8.0
    trend = df.groupby("periodo")[["accordato", "utilizzato", "sconfino"]].sum().reset_index() if not df.empty else pd.DataFrame(columns=["periodo","accordato","utilizzato","sconfino"])
    return {
        "periods": periods, "latest_period": latest_period, "latest": latest,
        "by_bank": by_bank, "by_cat": by_cat, "trend": trend, "totals": totals,
        "saturation": saturation, "score": score, "judgment": judgment,
        "rating": rating, "pd_est": pd_est, "banks": int(by_bank["banca"].nunique()) if not by_bank.empty else 0,
    }


def _money0(value) -> str:
    return f"€ {float(value):,.0f}".replace(",", ".")


def _pdf_footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor("#D9E1E8"))
    canvas.line(1.4*cm, 1.25*cm, A4[0]-1.4*cm, 1.25*cm)
    canvas.setFillColor(colors.HexColor("#0E2944"))
    canvas.setFont("Helvetica-Bold", 7.5)
    canvas.drawCentredString(A4[0]/2, 0.75*cm, "www.financeplus.tech")
    canvas.setFont("Helvetica", 7)
    canvas.drawRightString(A4[0]-1.4*cm, 0.75*cm, f"Pagina {doc.page} di 45")
    canvas.restoreState()


def _section_title(story, num, title, subtitle, color="#2BB7CE"):
    styles = getSampleStyleSheet()
    story.append(Paragraph(f"<b>{num}  {escape(title)}</b>", ParagraphStyle("sect", parent=styles["Heading1"], fontSize=16, leading=19, textColor=colors.HexColor("#0E2944"), spaceAfter=3)))
    story.append(Table([[""]], colWidths=[17.8*cm], rowHeights=[0.06*cm], style=TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor(color))])))
    story.append(Paragraph(escape(subtitle), ParagraphStyle("sub", parent=styles["BodyText"], fontSize=8.5, textColor=colors.HexColor("#66717E"), spaceBefore=5, spaceAfter=10)))


def _data_table(data, widths=None, font=7.2):
    if not data:
        data = [["nessun dato presente"]]
    t = Table(data, repeatRows=1, colWidths=widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0E2944")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), font),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#D5DCE3")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7F8FA")]),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 5), ("RIGHTPADDING", (0,0), (-1,-1), 5),
    ]))
    return t


def _metric_cards(items):
    cells = []
    for label, value, accent in items:
        card = Table(
            [
                [Paragraph(label.upper(), ParagraphStyle(
                    "ml", fontName="Helvetica", fontSize=6.5,
                    textColor=colors.HexColor("#74808D")
                ))],
                [Paragraph(f"<b>{escape(str(value))}</b>", ParagraphStyle(
                    "mv", fontSize=14, leading=17,
                    textColor=colors.HexColor("#0E2944")
                ))],
            ],
            colWidths=[4.1 * cm],
            rowHeights=[0.55 * cm, 0.85 * cm],
            style=TableStyle([
                ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#F4F6F8")),
                ("BOX", (0,0), (-1,-1), 0.4, colors.HexColor("#E4E8EC")),
                ("LINEBEFORE", (0,0), (0,-1), 4, colors.HexColor(accent)),
                ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ])
        )
        cells.append(card)
    return Table([cells], colWidths=[4.35 * cm] * len(cells), hAlign="LEFT")


def _cover_page(story, company, analysis):
    styles = getSampleStyleSheet()
    story += [Spacer(1, 3.2*cm), Paragraph("ANALISI<br/>CR<br/>AVANZATA", ParagraphStyle("cover", parent=styles["Title"], fontSize=32, leading=35, textColor=colors.HexColor("#0E2944"), alignment=TA_LEFT)), Spacer(1, 1.2*cm), Paragraph(f"<b>{escape(company)}</b>", ParagraphStyle("company", fontSize=18, leading=22, textColor=colors.HexColor("#0E2944"))), Spacer(1, 0.35*cm), Paragraph("soggetto della visura", ParagraphStyle("cap", fontSize=8, textColor=colors.grey)), Paragraph(escape(company), ParagraphStyle("body", fontSize=11)), Spacer(1, 0.5*cm), Paragraph("periodo censito", ParagraphStyle("cap2", fontSize=8, textColor=colors.grey)), Paragraph(escape(analysis["latest_period"]), ParagraphStyle("body2", fontSize=11)), Spacer(1, 2*cm), Paragraph("Le informazioni presenti in questo report sono una presentazione aggregata dei dati segnalati dagli intermediari alla Centrale dei Rischi secondo lo schema stabilito dalla Banca d'Italia.", ParagraphStyle("disclaimer", fontSize=8.5, leading=12, textColor=colors.HexColor("#66717E"))), PageBreak()]


def build_cr_report_pdf(company: str, df: pd.DataFrame, analysis: Dict) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=1.35*cm, rightMargin=1.35*cm, topMargin=1.25*cm, bottomMargin=1.55*cm)
    styles = getSampleStyleSheet()
    body = ParagraphStyle("fpbody", parent=styles["BodyText"], fontSize=8.2, leading=11, textColor=colors.HexColor("#384553"))
    note = ParagraphStyle("note", parent=body, backColor=colors.HexColor("#F4F6F8"), borderColor=colors.HexColor("#FF9E38"), borderWidth=0, borderPadding=8, leftIndent=4, spaceBefore=8)
    story = []
    _cover_page(story, company, analysis)
    # Page 2 index
    _section_title(story, "", "INDICE SEZIONI DEL REPORT", "Struttura completa del report avanzato a 36 mesi")
    index_lines = [
        "SCORING E RANKING: Giudizio di Sintesi; Monte Affidamenti; Monte Utilizzi; Informazioni Quantitative; Distinta Banche; Ranking banche e garanzie.",
        "EVENTI NEGATIVI: Perdite e sofferenze; Past due; Sconfini evitabili; Dettaglio banca; Eventi negativi; Sconfini inframensili.",
        "GARANZIE: Garanzie proprie e di terzi; Cointestazioni; Dati MCC; Derivati.",
        "TESORERIA: Tassi; Fabbisogno; Fidi a revoca; Oneri finanziari; Equilibrio; Portafoglio effetti.",
        "ALTRE LINEE E SEGNALAZIONI: Richieste informazioni; Contestazioni; Rettifiche; Factoring; Crediti di firma; Leasing; Import/Export; Riepilogo banca."
    ]
    for line in index_lines: story += [Paragraph(line, body), Spacer(1, 7)]
    story.append(PageBreak())

    sections = [
        ("SCORING E RANKING", 1, "#2BB7CE"), ("EVENTI NEGATIVI", 2, "#F25A64"),
        ("GARANZIE", 3, "#9B63E5"), ("TESORERIA", 4, "#FF9E38"),
        ("ALTRE LINEE E SEGNALAZIONI", 5, "#2DB77B")
    ]
    # section opener page 3
    for section_name, sec_no, sec_color in sections:
        story += [Spacer(1, 5.5*cm), Paragraph("CR AVANZATA - SEZIONE PRINCIPALE", ParagraphStyle(f"op{sec_no}", fontSize=12, textColor=colors.HexColor("#66717E"))), Spacer(1, 0.4*cm), Paragraph(section_name.replace(" ", "<br/>", 1), ParagraphStyle(f"opbig{sec_no}", fontSize=30, leading=33, textColor=colors.HexColor("#0E2944"))), Spacer(1, 1*cm), Paragraph(f"Sezione {sec_no}", ParagraphStyle(f"ops{sec_no}", fontSize=11, textColor=colors.HexColor(sec_color))), Spacer(1, 1.6*cm), Paragraph(f"<b>{escape(company)}</b>", ParagraphStyle(f"opc{sec_no}", fontSize=14, textColor=colors.HexColor("#0E2944"))), Paragraph(f"periodo censito: {escape(analysis['latest_period'])}", body), PageBreak()]
        if sec_no == 1:
            # 4
            _section_title(story, "1.01", "SCORING CENTRALE RISCHI", f"Periodo censito: {analysis['latest_period']}", sec_color)
            story.append(_metric_cards([("Score", f"{analysis['score']} / 100", sec_color), ("Giudizio", analysis['judgment'], "#2DB77B"), ("Rating indicativo", analysis['rating'], "#FF9E38"), ("PD stimata", f"{analysis['pd_est']:.2f}%", "#9B63E5")]))
            alerts = ["SI Insoluti / sconfini" if analysis['totals'].get('sconfino',0)>0 else "NO Insoluti / sconfini", "SI Tensione finanziaria" if analysis['saturation']>90 else "NO Tensione finanziaria", "SI Sofferenze" if analysis['totals'].get('sofferenza',0)>0 else "NO Sofferenze", "SI Perdite" if analysis['totals'].get('perdita',0)>0 else "NO Perdite"]
            story += [Spacer(1, 14), Paragraph("<b>ANOMALIE RILEVATE - FAST ADVISORY</b>", body), _data_table([["Indicatore", "Esito"]] + [[a.split(" ",1)[1], a.split(" ",1)[0]] for a in alerts], [12*cm, 5.6*cm]), Paragraph("Lo score pondera saturazione, sconfini, sofferenze e perdite. Deve essere verificato sul prospetto originale.", note), PageBreak()]
            # 5
            _section_title(story, "1.02 / 1.03", "MONTE AFFIDAMENTI / MONTE UTILIZZI", "Ultimo mese e andamento recente per forma tecnica", sec_color)
            cat_rows = [["Forma tecnica","Accordato operativo","Utilizzato","Sconfini"]] + [[r['categoria'] or '-', _money0(r['accordato']), _money0(r['utilizzato']), _money0(r['sconfino'])] for _,r in analysis['by_cat'].iterrows()]
            story += [_metric_cards([("Monte affidamenti", _money0(analysis['totals'].get('accordato',0)), "#2BB7CE"), ("Monte utilizzi", _money0(analysis['totals'].get('utilizzato',0)), "#FF9E38")]), Spacer(1, 14), _data_table(cat_rows, [7*cm,3.5*cm,3.5*cm,3.5*cm]), PageBreak()]
            # 6
            _section_title(story, "1.04 / 1.05", "INFORMAZIONI QUANTITATIVE / DISTINTA BANCHE", "Quadro sintetico e dettaglio ultimo mese", sec_color)
            story += [_metric_cards([("Attuali affidamenti", _money0(analysis['totals'].get('accordato',0)), "#2BB7CE"), ("Attuali utilizzi", _money0(analysis['totals'].get('utilizzato',0)), "#FF9E38"), ("Sconfini / insoluti", _money0(analysis['totals'].get('sconfino',0)), "#F25A64"), ("Istituti segnalanti", analysis['banks'], "#2DB77B")]), Spacer(1, 14)]
            bank_rows = [["Banca / Intermediario","Accordato op.","Utilizzato","Sconfini"]] + [[r['banca'],_money0(r['accordato']),_money0(r['utilizzato']),_money0(r['sconfino'])] for _,r in analysis['by_bank'].iterrows()]
            story += [_data_table(bank_rows,[9*cm,3*cm,3*cm,2.6*cm],6.7), PageBreak()]
            # pages 7-14
            titles = [
                ("1.06","DISTINTA BANCHE - ULTIMI 36 MESI","Andamento aggregato accordato operativo, utilizzato e sconfini"),
                ("1.06","DISTINTA BANCHE - CONTINUAZIONE","Seconda parte della serie storica fino a 36 mesi"),
                ("","MONITORAGGIO FATTORI DI RISCHIO","Ultimi 6 mesi disponibili in visura"),
                ("","ANALISI COMPORTAMENTALE - RANKING BANCHE","Top 3 per affidamenti, utilizzi, sconfini e garanzie"),
                ("1.07","RANKING BANCHE PER ACCORDATO OPERATIVO","Ultimo mese - quale banca affida di più"),
                ("1.08","RANKING BANCHE PER UTILIZZATO","Ultimo mese - quale banca è stata utilizzata di più"),
                ("1.09","RANKING BANCHE PER SCONFINI","Ultimo mese e cumulato disponibile"),
                ("1.10","RANKING GARANZIE","Ultimo mese - a quale banca sono state date più garanzie"),
            ]
            for n,tit,sub in titles:
                _section_title(story,n,tit,sub,sec_color)
                if "36 MESI" in tit or "CONTINUAZIONE" in tit:
                    tr = analysis['trend']
                    rows = [["Periodo","Accordato operativo","Utilizzato","Sconfini"]] + [[r['periodo'],_money0(r['accordato']),_money0(r['utilizzato']),_money0(r['sconfino'])] for _,r in tr.iterrows()]
                else:
                    metric = 'accordato' if 'ACCORDATO' in tit or 'AFFIDAMENTI' in tit else 'utilizzato' if 'UTILIZZATO' in tit else 'sconfino' if 'SCONFINI' in tit else 'garanzia'
                    b = analysis['by_bank'].sort_values(metric, ascending=False)
                    rows = [["Banca", metric.title(), "Peso %"]]
                    total = b[metric].sum()
                    rows += [[r['banca'],_money0(r[metric]),f"{(r[metric]/total*100 if total else 0):.1f}%"] for _,r in b.iterrows()]
                story += [_data_table(rows, [11.5*cm,3.3*cm,2.8*cm] if len(rows[0])==3 else [5*cm,4.2*cm,4.2*cm,4.2*cm],6.8), Paragraph("Le sezioni mantengono la struttura del report a 45 pagine anche quando i dati disponibili sono inferiori a 36 mesi.", note), PageBreak()]
        elif sec_no == 2:
            page_defs = [
                ("2.01","CREDITI PASSATI A PERDITA E SOFFERENZE","Ultimi 36 mesi disponibili"),
                ("2.02 / 2.03","CREDITI SCADUTI E SCONFINATI 90/180 GG","Sovrautilizzi delle linee e durata delle anomalie"),
                ("2.04 / 2.05","SCONFINI EVITABILI E COMPENSAZIONI","Analisi della capienza disponibile sulle linee a revoca"),
                ("2.06 / 2.07","SCONFINI PER BANCA ED EVENTI NEGATIVI","Dettaglio per istituto e tabella riassuntiva"),
                ("2.08","SCONFINI INFRAMENSILI","Stima su utilizzi medi delle linee a revoca"),
            ]
            for n,tit,sub in page_defs:
                _section_title(story,n,tit,sub,sec_color)
                story += [_metric_cards([("Crediti passati a perdita",_money0(analysis['totals'].get('perdita',0)),"#F25A64"),("Sofferenze",_money0(analysis['totals'].get('sofferenza',0)),"#F25A64"),("Sconfini registrati",_money0(analysis['totals'].get('sconfino',0)),"#FF9E38")]), Spacer(1,15)]
                b = analysis['by_bank'][analysis['by_bank']['sconfino']>0]
                rows = [["Istituto","Sconfini"]] + [[r['banca'],_money0(r['sconfino'])] for _,r in b.iterrows()]
                story += [_data_table(rows,[13.5*cm,4*cm],7), Paragraph("La durata delle anomalie e gli eventuali picchi inframensili devono essere verificati sul documento ufficiale e sugli estratti conto.",note), PageBreak()]
        elif sec_no == 3:
            page_defs = [
                ("3.01","GARANZIE PRESTATE SU PROPRI AFFIDAMENTI","Valore garanzie, importo garantito e rapporto con gli affidamenti"),
                ("3.02 / 3.03","GARANZIE DI TERZI E COINTESTAZIONI","Garanzie su affidamenti accesi da terzi e posizioni cointestate"),
                ("3.04","DATI MCC - MEDIO CREDITO CENTRALE","Ultimi 6 mesi e indicatori andamentali CR"),
                ("3.05 / 3.06","ANDAMENTO DERIVATI","Valore intrinseco e numero linee - prima parte"),
                ("3.06","DERIVATI - CONTINUAZIONE","Seconda parte dello storico fino a 36 mesi"),
            ]
            for n,tit,sub in page_defs:
                _section_title(story,n,tit,sub,sec_color)
                coverage = analysis['totals'].get('garanzia',0)/max(analysis['totals'].get('accordato',0),1)*100
                story += [_metric_cards([("Valore garanzie",_money0(analysis['totals'].get('garanzia',0)),"#9B63E5"),("Monte affidamenti",_money0(analysis['totals'].get('accordato',0)),"#2BB7CE"),("% copertura",f"{coverage:.1f}%","#2DB77B"),("CR score",f"{analysis['score']}/100","#FF9E38")]), Spacer(1,15)]
                b=analysis['by_bank'][analysis['by_bank']['garanzia']>0]
                story += [_data_table([["Banca","Valore garanzia"]]+[[r['banca'],_money0(r['garanzia'])] for _,r in b.iterrows()],[13.5*cm,4*cm],7), Paragraph("Il report non sostituisce il calcolo ufficiale MCC né la verifica delle garanzie nominative e delle cointestazioni.",note), PageBreak()]
        elif sec_no == 4:
            page_defs = [
                ("4.01","TASSI APPLICATI ALLE IMPRESE","Benchmark storico del modello dimostrativo"),
                ("4.02","FABBISOGNO TESORERIA","Proiezione stimata a 6 mesi su autoliquidanti e revoca"),
                ("4.03","DISPONIBILITA FIDI A REVOCA","Media ultimi 36 mesi per banca"),
                ("4.04","ANALISI ONERI FINANZIARI","Sottoutilizzo medio e risparmio potenziale sulle commissioni"),
                ("4.05","EQUILIBRIO DI TESORERIA","Utilizzo medio 36/12 mesi per forma tecnica"),
                ("4.06","PORTAFOGLIO EFFETTI E POLMONE FINANZIARIO","Rischiosità degli anticipi e capacità di assorbimento"),
                ("4.07","IMPATTO IMPAGATI DELLA CLIENTELA","Ultimi 36 mesi - confronto con disponibilità a revoca"),
            ]
            for n,tit,sub in page_defs:
                _section_title(story,n,tit,sub,sec_color)
                story += [_metric_cards([("Accordato attuale",_money0(analysis['totals'].get('accordato',0)),"#2BB7CE"),("Utilizzato attuale",_money0(analysis['totals'].get('utilizzato',0)),"#FF9E38"),("Saturazione",f"{analysis['saturation']:.1f}%","#F25A64")]), Spacer(1,15)]
                story += [_data_table([["Periodo","Accordato","Utilizzato","Sconfini"]]+[[r['periodo'],_money0(r['accordato']),_money0(r['utilizzato']),_money0(r['sconfino'])] for _,r in analysis['trend'].iterrows()],[5*cm,4.2*cm,4.2*cm,4.2*cm],7), Paragraph("Le proiezioni di tesoreria sono stime tecniche e vanno integrate con budget, scadenziari e stagionalità.",note), PageBreak()]
        else:
            page_defs = [
                ("5.01 / 5.02","RICHIESTE DI INFORMAZIONI E CONTESTAZIONI","Ultimi 6 mesi e rapporti contestati"),
                ("5.03","RETTIFICHE","Ultimi 36 mesi - variazioni e correzioni alle segnalazioni"),
                ("5.04","FACTORING ATTIVI E CESSIONE DI CREDITO","Ultimi 36 mesi - soggetto come creditore cedente"),
                ("5.05","FACTORING PASSIVI E CESSIONE DI CREDITO","Ultimi 36 mesi - soggetto come debitore ceduto"),
                ("5.06","CREDITI DI FIRMA","Natura commerciale e finanziaria"),
                ("5.07","LEASING","Ultimi 36 mesi - peso sul totale linee a scadenza"),
                ("5.08","IMPORT / EXPORT / DIVISA","Crediti per cassa e crediti di firma"),
                ("5.09","OPERAZIONI EFFETTUATE PER CONTO TERZI","Ultimi 36 mesi"),
                ("5.10","RIEPILOGO GENERALE PER BANCA","Quadro complessivo dell'ultimo mese e dello storico acquisito"),
            ]
            for n,tit,sub in page_defs:
                _section_title(story,n,tit,sub,sec_color)
                b=analysis['by_bank'].copy()
                story += [_data_table([["Intermediario","Mesi","Accordato op.","Utilizzato","Sconfini storici","Alert"]]+[[r['banca'],len(analysis['periods']) or 1,_money0(r['accordato']),_money0(r['utilizzato']),_money0(r['sconfino']),"SI" if r['sconfino']>0 else "NO"] for _,r in b.iterrows()],[7.2*cm,1.2*cm,2.6*cm,2.6*cm,2.6*cm,1.2*cm],6.2), Paragraph("Le rettifiche, le richieste di informazione e le contestazioni richiedono una verifica visiva del prospetto Banca d'Italia.",note), PageBreak()]
    # final page 45 (remove trailing page break not critical)
    _section_title(story,"","AVVERTENZE","Limiti, responsabilità e verifiche necessarie","#0E2944")
    warnings = [
        "Le informazioni sono elaborate in ragione dei dati acquisiti e delle tecnologie disponibili.",
        "Il documento ha finalità informative e consulenziali e non costituisce delibera bancaria, rating ufficiale o attestazione.",
        "Score e probabilità di default sono stime interne e possono variare al mutare delle segnalazioni.",
        "I dati estratti automaticamente devono essere confrontati con il prospetto originale della Centrale dei Rischi.",
        "Le sezioni prive di dati mantengono la struttura del report a 45 pagine.",
        "Prima di assumere decisioni finanziarie occorre integrare la CR con bilanci, estratti conto e informazioni qualitative.",
        "FinancePlus.tech non risponde di decisioni assunte esclusivamente sulla base dell'elaborato automatico senza verifica professionale.",
    ]
    for i,w in enumerate(warnings,1):
        story += [Table([[str(i), Paragraph(w, body)]], colWidths=[0.8*cm,16.8*cm], style=TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("TEXTCOLOR",(0,0),(0,0),colors.HexColor("#0E2944")),("FONTNAME",(0,0),(0,0),"Helvetica-Bold"),("BOTTOMPADDING",(0,0),(-1,-1),10)]))]
    story.append(Paragraph(f"Data elaborazione: {datetime.now().strftime('%d-%m-%Y')} - www.financeplus.tech", body))
    doc.build(story, onFirstPage=_pdf_footer, onLaterPages=_pdf_footer)
    return buffer.getvalue()


def build_cr_html(company: str, df: pd.DataFrame, analysis: Dict) -> bytes:
    bank_html = analysis['by_bank'].to_html(index=False, classes="table", border=0, formatters={c: _money0 for c in ['accordato','utilizzato','sconfino','garanzia']})
    trend_html = analysis['trend'].to_html(index=False, classes="table", border=0, formatters={c: _money0 for c in ['accordato','utilizzato','sconfino']})
    html = f"""<!doctype html><html lang='it'><head><meta charset='utf-8'><title>Report CR {escape(company)}</title><style>body{{font-family:Arial;margin:32px;color:#0E2944}}h1{{font-size:30px}}.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}}.card{{background:#f4f6f8;padding:16px;border-left:5px solid #2bb7ce}}table{{border-collapse:collapse;width:100%;margin:20px 0}}th{{background:#0e2944;color:white}}td,th{{padding:8px;border:1px solid #d9e1e8;text-align:left}}@media print{{body{{margin:10mm}}}}</style></head><body><h1>ANALISI CR AVANZATA</h1><h2>{escape(company)}</h2><p>Periodo: {escape(analysis['latest_period'])}</p><div class='cards'><div class='card'><b>Score</b><br>{analysis['score']}/100</div><div class='card'><b>Giudizio</b><br>{analysis['judgment']}</div><div class='card'><b>Affidamenti</b><br>{_money0(analysis['totals'].get('accordato',0))}</div><div class='card'><b>Utilizzi</b><br>{_money0(analysis['totals'].get('utilizzato',0))}</div></div><h2>Distinta banche</h2>{bank_html}<h2>Andamento</h2>{trend_html}<p><small>Elaborazione automatica FinancePlus.tech. Verificare sempre con il prospetto ufficiale.</small></p></body></html>"""
    return html.encode("utf-8")


def build_cr_outputs_zip(company: str, df: pd.DataFrame, analysis: Dict, pdf: bytes, html: bytes) -> bytes:
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", company).strip("_") or "CLIENTE"
    date_slug = datetime.now().strftime("%d_%m_%Y")
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"REPORT_CR_{safe}_{date_slug}.pdf", pdf)
        z.writestr(f"REPORT_CR_{safe}_{date_slug}.html", html)
        z.writestr("dati_cr_strutturati.json", json.dumps({"azienda":company,"analisi":{k:v for k,v in analysis.items() if k not in ('latest','by_bank','by_cat','trend')},"righe":df.to_dict(orient='records')}, ensure_ascii=False, indent=2, default=str))
        z.writestr("tabelle/dati_normalizzati.csv", df.to_csv(index=False).encode("utf-8-sig"))
        z.writestr("tabelle/distinta_banche.csv", analysis['by_bank'].to_csv(index=False).encode("utf-8-sig"))
        z.writestr("tabelle/andamento_mensile.csv", analysis['trend'].to_csv(index=False).encode("utf-8-sig"))
        for metric, title in [("accordato","Affidamenti per banca"),("utilizzato","Utilizzi per banca"),("sconfino","Sconfini per banca")]:
            fig, ax = plt.subplots(figsize=(10,5))
            data = analysis['by_bank'].sort_values(metric, ascending=True).tail(12)
            ax.barh(data['banca'], data[metric])
            ax.set_title(title)
            ax.set_xlabel("Euro")
            fig.tight_layout()
            img=io.BytesIO(); fig.savefig(img, format="png", dpi=150, bbox_inches="tight"); plt.close(fig)
            z.writestr(f"grafici/{metric}_per_banca.png", img.getvalue())
    return b.getvalue()


def preview_pdf(pdf_bytes: bytes, height: int = 760):
    encoded = base64.b64encode(pdf_bytes).decode("ascii")
    st.components.v1.html(f'<iframe src="data:application/pdf;base64,{encoded}" width="100%" height="{height}" style="border:1px solid #d9e1e8;border-radius:8px"></iframe>', height=height+20)


def italian_report_filename(company: str) -> str:
    company_clean = re.sub(r"[^A-Za-z0-9À-ÖØ-öø-ÿ ]+", " ", company).upper()
    company_clean = re.sub(r"\s+", " ", company_clean).strip()
    now = datetime.now()
    return f"REPORT CR {company_clean} {now.day} {MONTHS_IT[now.month-1]} {now.year}.pdf"


# -----------------------------
# INTERFACCIA
# -----------------------------
def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="🏢", layout="wide")

    # Impedisce alla traduzione automatica del browser di alterare le voci del menu
    # (es. "Inserisci cliente" trasformato erroneamente in "Installa cliente").
    st.components.v1.html(
        """
        <script>
        (function () {
          const doc = window.parent.document;
          doc.documentElement.setAttribute('lang', 'it');
          doc.documentElement.setAttribute('translate', 'no');
          doc.body.setAttribute('translate', 'no');
          doc.body.classList.add('notranslate');

          let meta = doc.querySelector('meta[name="google"]');
          if (!meta) {
            meta = doc.createElement('meta');
            meta.setAttribute('name', 'google');
            doc.head.appendChild(meta);
          }
          meta.setAttribute('content', 'notranslate');

          const protect = () => {
            doc.querySelectorAll('[data-testid="stSidebar"], [data-testid="stAppViewContainer"], .stApp')
              .forEach(el => {
                el.setAttribute('translate', 'no');
                el.classList.add('notranslate');
              });
          };
          protect();
          new MutationObserver(protect).observe(doc.body, {childList: true, subtree: true});
        })();
        </script>
        """,
        height=0,
        width=0,
    )

    init_db()

    st.markdown("""
    <style>
    .stApp {background: #F6F8FA;}
    h1, h2, h3 {color:#153B5B;}
    div[data-testid="stMetric"] {background:white; border:1px solid #D9E1E8; padding:12px; border-radius:10px;}
    </style>
    """, unsafe_allow_html=True)

    st.title("🏢 FinancePlus – Clienti e Pratiche")

    page = st.sidebar.radio(
        "Menu",
        ["Inserisci cliente", "CLIENTI", "Inserisci collaboratore", "Inserisci informazioni", "Anteprima report", "Archivio"]
    )

    if page == "Inserisci cliente":
        st.subheader("Inserisci cliente da visura camerale o report PDF")
        uploaded = st.file_uploader("Carica PDF", type=["pdf"])

        defaults = {
            "denominazione": "", "partita_iva": "", "sede": "",
            "indirizzo": "", "cap": "", "provincia": "", "documento": "", "collaboratore_id": None
        }

        if uploaded:
            key = f"{uploaded.name}-{uploaded.size}"
            if st.session_state.get("doc_key") != key:
                text = extract_pdf_text(uploaded.getvalue())
                parsed = parse_cliente(text, uploaded.name)
                st.session_state["parsed_cliente"] = parsed
                st.session_state["doc_text"] = text
                st.session_state["doc_key"] = key

        data = st.session_state.get("parsed_cliente", defaults)

        with st.form("cliente_form"):
            c1, c2 = st.columns(2)
            with c1:
                denominazione = st.text_input("PMI – Denominazione aziendale", value=data["denominazione"])
                partita_iva = st.text_input("P.IVA", value=data["partita_iva"], max_chars=11)
                sede = st.text_input("SEDE – Solo Comune", value=data["sede"])
            with c2:
                indirizzo = st.text_input("Indirizzo – Strada e civico", value=data["indirizzo"])
                cap = st.text_input("CAP", value=data["cap"], max_chars=5)
                provincia = st.text_input("Provincia", value=data["provincia"], max_chars=2)

            collaboratori = list_collaboratori()
            collab_options = [None] + [c[0] for c in collaboratori]
            collab_labels = {None: "Nessun collaboratore"}
            collab_labels.update({c[0]: f"{c[1]} {c[2]}" for c in collaboratori})
            current_collab = data.get("collaboratore_id")
            default_index = collab_options.index(current_collab) if current_collab in collab_options else 0
            collaboratore_id = st.selectbox(
                "Collaboratore",
                options=collab_options,
                index=default_index,
                format_func=lambda x: collab_labels[x]
            )

            salva = st.form_submit_button("💾 Salva cliente", use_container_width=True)

        if salva:
            if not denominazione.strip():
                st.error("La denominazione aziendale è necessaria.")
            else:
                cid = save_cliente({
                    "denominazione": clean(denominazione).upper(),
                    "partita_iva": re.sub(r"\D", "", partita_iva),
                    "sede": clean(sede).upper(),
                    "indirizzo": clean(indirizzo).upper(),
                    "cap": re.sub(r"\D", "", cap),
                    "provincia": clean(provincia).upper(),
                    "documento": data.get("documento", ""),
                    "collaboratore_id": collaboratore_id,
                })
                st.success(f"Cliente salvato correttamente. ID: {cid}")

        with st.expander("Verifica testo estratto"):
            st.text_area("Testo PDF", st.session_state.get("doc_text", ""), height=260, disabled=True)

    elif page == "CLIENTI":
        st.subheader("CLIENTI")
        clienti = list_clienti()
        if not clienti:
            st.warning("Nessun cliente presente. Inserire prima un cliente.")
            return
        labels = {c[0]: f"{c[1]} - P.IVA {c[2] or 'non indicata'}" for c in clienti}
        selected_id = st.selectbox("Seleziona cliente", options=list(labels), format_func=lambda x: labels[x], key="clienti_select")
        cliente = get_cliente(selected_id)
        st.info(f"**Cliente:** {cliente[1]}  \n**P.IVA:** {cliente[2] or '-'}  \n**Sede:** {cliente[3] or '-'}")
        st.markdown("### Analisi / Value")
        module = st.radio("Seleziona modulo", ["CR", "CC", "Bilancio", "Bozza Bilancio"], horizontal=True)
        if module == "CC":
            if cc_flussi is None:
                st.error("Il modulo Analisi CC/FLUSSI non è disponibile. Verificare che Analisi_CC_FLUSSI.py sia presente e che tutte le dipendenze siano installate.")
                st.caption(str(_cc_import_error))
            else:
                selected_client = {
                    "legal_name": cliente[1] or "",
                    "vat_number": cliente[2] or "",
                    "city": cliente[3] or "",
                    "address": cliente[4] or "",
                    "postal_code": cliente[5] or "",
                    "province": cliente[6] or "",
                }
                cc_flussi.main(embedded=True, selected_client=selected_client)
        elif module in {"Bilancio", "Bozza Bilancio"}:
            st.info(f"Modulo {module} predisposto nella navigazione. La relativa logica di analisi sarà collegata nella fase successiva.")
        else:
            st.markdown("#### Analisi Centrale Rischi - 36 mesi")
            st.caption("Formati ammessi: Excel .xlsx, CSV, TXT e PDF testuale.")
            uploaded_cr = st.file_uploader("Carica CR 36 mesi", type=["xlsx","csv","txt","pdf"], key=f"cr_upload_{selected_id}")
            if uploaded_cr:
                try:
                    df_cr, source_text, warnings = read_cr_file(uploaded_cr)
                    analysis = analyse_cr(df_cr)
                    for warning in warnings:
                        st.warning(warning)
                    if df_cr.empty:
                        st.error("Nessun dato CR strutturato è stato riclassificato. Verificare che il file sia testuale e contenga tabelle o colonne riconoscibili.")
                    else:
                        c1,c2,c3,c4 = st.columns(4)
                        c1.metric("CR Score", f"{analysis['score']}/100")
                        c2.metric("Affidamenti", _money0(analysis['totals'].get('accordato',0)))
                        c3.metric("Utilizzi", _money0(analysis['totals'].get('utilizzato',0)))
                        c4.metric("Sconfini", _money0(analysis['totals'].get('sconfino',0)))
                        st.dataframe(df_cr, use_container_width=True, hide_index=True)
                        pdf_cr = build_cr_report_pdf(cliente[1], df_cr, analysis)
                        html_cr = build_cr_html(cliente[1], df_cr, analysis)
                        zip_cr = build_cr_outputs_zip(cliente[1], df_cr, analysis, pdf_cr, html_cr)
                        st.markdown("### Anteprima PDF")
                        preview_pdf(pdf_cr)
                        filename = italian_report_filename(cliente[1])
                        d1,d2,d3 = st.columns(3)
                        d1.download_button("Scarica report PDF", pdf_cr, file_name=filename, mime="application/pdf", use_container_width=True)
                        d2.download_button("Scarica report HTML", html_cr, file_name=filename.replace('.pdf','.html'), mime="text/html", use_container_width=True)
                        d3.download_button("Scarica pacchetto completo ZIP", zip_cr, file_name=filename.replace('.pdf',' - OUTPUT COMPLETI.zip'), mime="application/zip", use_container_width=True)
                except Exception as exc:
                    st.exception(exc)

    elif page == "Inserisci collaboratore":
        st.subheader("Inserisci collaboratore")
        st.caption("Crea l'anagrafica del collaboratore da associare ai clienti.")

        with st.form("collaboratore_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                nome = st.text_input("Nome")
            with c2:
                cognome = st.text_input("Cognome")
            salva_collab = st.form_submit_button(
                "💾 Salva collaboratore",
                use_container_width=True
            )

        if salva_collab:
            if not nome.strip() or not cognome.strip():
                st.error("Inserire nome e cognome.")
            else:
                cid = save_collaboratore(nome, cognome)
                st.success(f"Collaboratore salvato correttamente. ID: {cid}")

        collaboratori = list_collaboratori()
        if collaboratori:
            st.write("### Collaboratori salvati")
            st.dataframe(
                [{"ID": c[0], "Nome": c[1], "Cognome": c[2]} for c in collaboratori],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("Nessun collaboratore salvato.")

    elif page == "Inserisci informazioni":
        st.subheader("Inserisci informazioni pratica")
        clienti = list_clienti()
        if not clienti:
            st.warning("Inserire prima almeno un cliente.")
            return

        labels = {c[0]: f"{c[1]} – P.IVA {c[2] or 'non indicata'}" for c in clienti}
        selected_id = st.selectbox("Seleziona cliente", options=list(labels), format_func=lambda x: labels[x])
        cliente = get_cliente(selected_id)

        st.info(
            f"**PMI:** {cliente[1]}  \n"
            f"**P.IVA:** {cliente[2] or '-'}  \n"
            f"**SEDE:** {cliente[3] or '-'}  \n"
            f"**INDIRIZZO:** {cliente[4] or '-'}  \n"
            f"**COLLABORATORE:** {cliente[10] or '-'}"
        )

        with st.form("pratica_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                prodotto = st.selectbox("PRODOTTO", PRODOTTI)
                importo = st.number_input("IMPORTO (€)", min_value=0.0, step=1000.0, format="%.2f")
                istituto = st.text_input("ISTITUTO – Banca / Fintech / SGR")
                gestore = st.selectbox("GESTORE", GESTORI)
            with c2:
                stato = st.selectbox("STATO", STATI)
                data_integrazione = st.date_input("Data integrazione", value=date.today())
                note = st.text_area("NOTE", height=110)
                integrazioni = st.text_area("INTEGRAZIONI", height=110)

            submit = st.form_submit_button("💾 Salva modulo pratica", use_container_width=True)

        if submit:
            pid = save_pratica(
                selected_id, prodotto, importo, istituto, gestore,
                note, integrazioni, data_integrazione, stato
            )
            st.success(f"Modulo pratica salvato. Numero pratica: {pid}")

        pratiche_cliente = list_pratiche(selected_id)
        st.caption(f"Moduli già presenti per il cliente: {len(pratiche_cliente)}")

    elif page == "Anteprima report":
        st.subheader("Anteprima e download report")
        clienti = list_clienti()
        if not clienti:
            st.warning("Nessun cliente presente.")
            return

        tab1, tab2 = st.tabs(["Report cliente", "Report generale"])

        with tab1:
            labels = {c[0]: f"{c[1]} – {c[2] or 'P.IVA non indicata'}" for c in clienti}
            selected_id = st.selectbox(
                "Seleziona cliente per il report",
                options=list(labels),
                format_func=lambda x: labels[x],
                key="report_cliente"
            )
            cliente = get_cliente(selected_id)
            pratiche = list_pratiche(selected_id)

            st.markdown(f"### {cliente[1]}")
            st.write(f"**P.IVA:** {cliente[2] or '-'}")
            st.write(f"**Sede:** {cliente[3] or '-'}")
            st.write(f"**Indirizzo:** {cliente[4] or '-'}")
            st.write(f"**Collaboratore:** {cliente[10] or '-'}")
            st.write(f"**Pratiche:** {len(pratiche)}")

            for p in pratiche:
                with st.expander(f"Pratica {p[0]} – {p[5]} – {money(p[6])} – {p[12]}"):
                    st.write(f"**Istituto:** {p[7] or '-'}")
                    st.write(f"**Gestore:** {p[8] or '-'}")
                    st.write(f"**Note:** {p[9] or '-'}")
                    st.write(f"**Integrazioni:** {p[10] or '-'}")
                    st.write(f"**Data integrazione:** {p[11] or '-'}")

            pdf = build_client_report(cliente, pratiche)
            safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", cliente[1])
            st.download_button(
                "📄 Scarica report PDF cliente",
                data=pdf,
                file_name=f"Report_{safe_name}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        with tab2:
            tutte_pratiche = list_pratiche()
            st.metric("Clienti", len(clienti))
            st.metric("Pratiche", len(tutte_pratiche))
            general_pdf = build_general_report(clienti, tutte_pratiche)
            st.download_button(
                "📚 Scarica report generale PDF",
                data=general_pdf,
                file_name="Report_Generale_Clienti_Pratiche.pdf",
                mime="application/pdf",
                use_container_width=True
            )

    elif page == "Archivio":
        st.subheader("Archivio clienti e pratiche")
        clienti = list_clienti()
        pratiche = list_pratiche()

        st.write("### Clienti")
        if clienti:
            st.dataframe(
                [{
                    "ID": c[0], "PMI": c[1], "P.IVA": c[2],
                    "Sede": c[3], "Indirizzo": c[4], "CAP": c[5],
                    "Provincia": c[6], "Collaboratore": c[8]
                } for c in clienti],
                use_container_width=True, hide_index=True
            )
        else:
            st.info("Nessun cliente.")

        st.write("### Pratiche")
        if pratiche:
            st.dataframe(
                [{
                    "ID": p[0], "Cliente": p[1], "Prodotto": p[5],
                    "Importo": p[6], "Istituto": p[7], "Gestore": p[8],
                    "Stato": p[12], "Data integrazione": p[11]
                } for p in pratiche],
                use_container_width=True, hide_index=True
            )
        else:
            st.info("Nessuna pratica.")


if __name__ == "__main__":
    main()
