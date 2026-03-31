"""
pratiche.py — Navigatore Pratiche Immigrazione
Sistema di estrazione via bookmarklet (senza Selenium).

Requisiti:
    pip install customtkinter openpyxl

Niente geckodriver, niente browser controllato.
"""

import customtkinter as ctk
from tkinter import messagebox, filedialog
import tkinter as tk
from tkinter import ttk
import threading
import json
import os
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
#  CONFIGURAZIONE
# ─────────────────────────────────────────────
try:
    OPERATORE = os.getlogin()
except Exception:
    OPERATORE = os.environ.get("USERNAME", "OPERATORE")

EXCEL_PATH_DEFAULT = f"pratiche_{OPERATORE}.xlsx"
SERVER_PORT        = 7432   # porta locale per ricevere dati dal bookmarklet

# ─────────────────────────────────────────────
#  COLONNE EXCEL
# ─────────────────────────────────────────────
COL_PRATICHE = [
    "Nr. Pratica", "Operatore Gestionale", "Stato Pratica",
    "Tipo Soggiorno", "Tipo Pratica", "Validita Soggiorno",
    "Cognome", "Nome", "CUI", "Sesso", "Data Nascita",
    "Luogo Nascita", "Nazione Nascita", "Cittadinanza",
    "Codice Fiscale", "Stato Civile", "Telefono",
    "Comune", "Indirizzo", "Motivo Soggiorno",
    "Coniuge", "Referenze", "Note Gestionale",
    "Data Presentazione", "Scadenza Rinnovo",
    "Estratto Da", "Data Estrazione",
]

COL_ATTIVITA = [
    "Nr. Pratica", "Stato",
    "SDI Esito", "SDI Note",
    "Reddito Tipo", "Reddito Note",
    "Famiglia Trainante", "Nome Trainante", "CF Trainante", "Nr Pratica Trainante",
    "Documenti Mancanti", "Note Personali",
    "Checklist Compilata Da", "Checklist Compilata Il", "Note Modificate Il",
]

STATI = {
    "da_verificare": ("🟡 Da verificare", "#FDE68A", "#B45309"),
    "sospesa":       ("🟠 Sospesa",       "#FCD34D", "#C2410C"),
    "validata":      ("🟢 Validata",      "#86EFAC", "#15803D"),
    "negata":        ("🔴 Negata",        "#FCA5A5", "#B91C1C"),
}

SDI_ESITI    = ["", "Positivo", "Negativo", "Positivo non ostativo"]
REDDITO_TIPI = ["", "Lavoro subordinato", "Lavoro autonomo", "Nessun reddito", "Altro"]

THEME = {
    "blu_scuro":   "#1F4E79",
    "blu_medio":   "#2E75B6",
    "sfondo":      "#F0F4F8",
    "card":        "#FFFFFF",
    "testo":       "#1E293B",
    "testo_light": "#64748B",
    "bordo":       "#E2E8F0",
    "verde":       "#16A34A",
    "verde_dark":  "#15803D",
    "arancio":     "#EA580C",
}

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ─────────────────────────────────────────────
#  BOOKMARKLET JavaScript
# ─────────────────────────────────────────────
BOOKMARKLET_JS = """javascript:(function(){{
  function vt(t){{
    var e=document.querySelector("input[title='"+t+"'],textarea[title='"+t+"']");
    return e?(e.value||e.innerText||"").trim():"";
  }}
  function vi(id){{
    var e=document.getElementById(id);
    return e?(e.value||e.innerText||"").trim():"";
  }}
  function vl(lbl){{
    var labels=Array.from(document.querySelectorAll("label"));
    var l=labels.find(function(x){{
      return x.innerText.trim().replace(":","").trim().toLowerCase()===lbl.toLowerCase();
    }});
    if(!l) return "";
    var td=l.closest("td");
    if(!td) return "";
    var next=td.nextElementSibling;
    if(!next) return "";
    var inp=next.querySelector("input");
    return inp?(inp.value||"").trim():"";
  }}
  var h3=document.querySelector("h3");
  var h3t=h3?h3.innerText:"";
  var m1=h3t.match(/Pratica\\s+n[°o]\\s*(\\S+)/i);
  var m2=h3t.match(/assegnata all'utente\\s+(\\S+)/i);
  var ne=document.getElementById("idnotexx");
  var dati={{
    "Nr. Pratica":          m1?m1[1]:"",
    "Operatore Gestionale": m2?m2[1]:"",
    "Stato Pratica":        vi("statoPraticaDescrConv"),
    "Tipo Soggiorno":       vi("docSoggiorno"),
    "Tipo Pratica":         vl("Tipo Pratica"),
    "Validita Soggiorno":   vt("Validit\\u00e0 del soggiorno"),
    "Cognome":              vl("Cognome"),
    "Nome":                 vl("Nome"),
    "CUI":                  vt("cui"),
    "Sesso":                vl("Sesso"),
    "Data Nascita":         vi("dataNascitaStraniero"),
    "Luogo Nascita":        vt("Luogo di Nascita"),
    "Nazione Nascita":      vl("Nazione Nascita"),
    "Cittadinanza":         vt("Cittadinanza dello straniero"),
    "Codice Fiscale":       vt("Codice fiscale dello straniero"),
    "Stato Civile":         vt("Stato Civile"),
    "Telefono":             vi("telefono"),
    "Comune":               vt("Comune di residenza in Italia"),
    "Indirizzo":            vt("Indirizzo"),
    "Motivo Soggiorno":     vi("motivoSoggiorno"),
    "Coniuge":              vt("Cognome e\\/o Nome del coniuge"),
    "Referenze":            vt("Eventuali referenze"),
    "Note Gestionale":      ne?(ne.value||ne.innerText||"").trim():"",
    "Data Presentazione":   vt("data di presentazione della istanza"),
    "Scadenza Rinnovo":     vt("Data scadenza rinnovo")
  }};
  if(!dati["Nr. Pratica"]){{
    alert("Nessuna pratica trovata.\\nAssicurati di essere sulla pagina della pratica.");
    return;
  }}
  fetch("http://localhost:{port}/pratica",{{
    method:"POST",
    headers:{{"Content-Type":"application/json"}},
    body:JSON.stringify(dati)
  }}).then(function(r){{
    if(r.ok) alert("\\u2705 Pratica "+dati["Nr. Pratica"]+" inviata all\'app!");
    else alert("\\u274c Errore nell\'invio. L\'app \\u00e8 aperta?");
  }}).catch(function(){{
    alert("\\u274c Impossibile connettersi all\'app.\\nAssicurati che Pratiche.exe sia aperto.");
  }});
}})();""".format(port=SERVER_PORT)


# ─────────────────────────────────────────────
#  SERVER HTTP LOCALE
# ─────────────────────────────────────────────
class _Handler(BaseHTTPRequestHandler):
    """Riceve i dati dal bookmarklet via POST /pratica"""
    callback = None   # viene impostato dall'App

    def do_POST(self):
        if self.path == "/pratica":
            try:
                length = int(self.headers.get("Content-Length", 0))
                body   = self.rfile.read(length)
                dati   = json.loads(body.decode("utf-8"))
                # Risposta CORS per permettere la fetch dal browser
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(b'{"ok":true}')
                if _Handler.callback:
                    _Handler.callback(dati)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
        else:
            self.send_response(404)
            self.end_headers()

    def do_OPTIONS(self):
        # Preflight CORS
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def log_message(self, *args):
        pass   # silenzia i log HTTP nel terminale


class ServerLocale:
    def __init__(self, callback):
        _Handler.callback = callback
        self._server = HTTPServer(("localhost", SERVER_PORT), _Handler)
        self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)

    def avvia(self):
        self._thread.start()

    def ferma(self):
        self._server.shutdown()


# ─────────────────────────────────────────────
#  LOGICA EXCEL
# ─────────────────────────────────────────────
def _bordo():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def _intestazione(ws, colonne, colore="1F4E79"):
    for col, nome in enumerate(colonne, start=1):
        c = ws.cell(row=1, column=col, value=nome)
        c.font      = Font(bold=True, color="FFFFFF", name="Segoe UI", size=10)
        c.fill      = PatternFill("solid", start_color=colore)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = _bordo()
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"

def crea_o_apri_excel(path: str) -> openpyxl.Workbook:
    if os.path.exists(path):
        return openpyxl.load_workbook(path)
    wb   = openpyxl.Workbook()
    ws_p = wb.active
    ws_p.title = "Pratiche"
    _intestazione(ws_p, COL_PRATICHE, "1F4E79")
    for i, w in enumerate([14,20,14,28,18,14,16,16,12,6,12,14,16,16,18,14,14,16,28,22,20,22,30,14,14,16,18], 1):
        ws_p.column_dimensions[get_column_letter(i)].width = w
    ws_a = wb.create_sheet("Attivita")
    _intestazione(ws_a, COL_ATTIVITA, "1E3A5F")
    for i, w in enumerate([14,18,20,30,20,30,16,22,18,16,35,40,20,18,18], 1):
        ws_a.column_dimensions[get_column_letter(i)].width = w
    wb.save(path)
    return wb

def _trova_riga(ws, nr: str, col: int = 1):
    for row in ws.iter_rows(min_row=2):
        if str(row[col - 1].value or "").strip() == nr.strip():
            return row[0].row
    return None

def _scrivi_riga(ws, riga: int, colonne: list, dati: dict):
    fill = PatternFill("solid", start_color="EFF6FF") if riga % 2 == 0 else None
    for col, nome in enumerate(colonne, start=1):
        c = ws.cell(row=riga, column=col, value=dati.get(nome, ""))
        c.font      = Font(name="Segoe UI", size=10)
        c.alignment = Alignment(vertical="center")
        c.border    = _bordo()
        if fill:
            c.fill = fill

def salva_pratica_excel(path: str, dati: dict):
    wb  = crea_o_apri_excel(path)
    ws  = wb["Pratiche"]
    nr  = dati.get("Nr. Pratica", "").strip()
    riga = _trova_riga(ws, nr) or (ws.max_row + 1)
    _scrivi_riga(ws, riga, COL_PRATICHE, dati)
    ws_a = wb["Attivita"]
    if not _trova_riga(ws_a, nr):
        _scrivi_riga(ws_a, ws_a.max_row + 1, COL_ATTIVITA,
                     {"Nr. Pratica": nr, "Stato": "da_verificare"})
    wb.save(path)

def leggi_pratiche(path: str) -> list[dict]:
    if not os.path.exists(path):
        return []
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Pratiche" not in wb.sheetnames:
        return []
    ws  = wb["Pratiche"]
    hdr = [c.value for c in ws[1]]
    return [dict(zip(hdr, r)) for r in ws.iter_rows(min_row=2, values_only=True) if any(r)]

def leggi_attivita(path: str, nr: str) -> dict:
    if not os.path.exists(path):
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Attivita" not in wb.sheetnames:
        return {}
    ws  = wb["Attivita"]
    hdr = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(hdr, row))
        if str(d.get("Nr. Pratica", "")).strip() == nr.strip():
            return d
    return {}

def leggi_tutti_stati(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Attivita" not in wb.sheetnames:
        return {}
    ws  = wb["Attivita"]
    hdr = [c.value for c in ws[1]]
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        d  = dict(zip(hdr, row))
        nr = str(d.get("Nr. Pratica","")).strip()
        if nr:
            out[nr] = str(d.get("Stato","da_verificare") or "da_verificare")
    return out

def salva_attivita_excel(path: str, dati: dict):
    wb  = openpyxl.load_workbook(path)
    ws  = wb["Attivita"]
    hdr = [c.value for c in ws[1]]
    nr  = str(dati.get("Nr. Pratica", "")).strip()
    riga = _trova_riga(ws, nr) or (ws.max_row + 1)
    fill = PatternFill("solid", start_color="EFF6FF") if riga % 2 == 0 else None
    for col, nome in enumerate(hdr, start=1):
        c = ws.cell(row=riga, column=col, value=dati.get(nome, ""))
        c.font      = Font(name="Segoe UI", size=10)
        c.alignment = Alignment(vertical="center")
        c.border    = _bordo()
        if fill:
            c.fill = fill
    wb.save(path)


# ─────────────────────────────────────────────
#  POPUP ANTEPRIMA
# ─────────────────────────────────────────────
class PopupAnteprima(ctk.CTkToplevel):
    def __init__(self, parent, dati: dict, on_conferma):
        super().__init__(parent)
        self.title("Anteprima estrazione")
        self.geometry("460x380")
        self.resizable(False, False)
        self.grab_set()
        self.lift()
        self.focus_force()
        self._dati        = dati
        self._on_conferma = on_conferma
        self._build(dati)

    def _build(self, d):
        hdr = ctk.CTkFrame(self, fg_color=THEME["blu_scuro"],
                           corner_radius=0, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        nome = f"{d.get('Cognome','')} {d.get('Nome','')}".strip() or "—"
        ctk.CTkLabel(hdr, text=f"  📋  {nome}",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color="white").pack(side="left", padx=14)
        ctk.CTkLabel(hdr, text=d.get("Nr. Pratica",""),
                     font=ctk.CTkFont(family="Consolas", size=12),
                     text_color="#93C5FD").pack(side="right", padx=14)

        body = ctk.CTkFrame(self, fg_color=THEME["sfondo"], corner_radius=0)
        body.pack(fill="both", expand=True, padx=14, pady=12)

        campi = [
            ("CUI",              d.get("CUI")),
            ("Codice Fiscale",   d.get("Codice Fiscale")),
            ("Nazione",          d.get("Nazione Nascita")),
            ("Data Nascita",     d.get("Data Nascita")),
            ("Tipo Soggiorno",   d.get("Tipo Soggiorno")),
            ("Motivo",           d.get("Motivo Soggiorno")),
            ("Comune",           d.get("Comune")),
            ("Scadenza Rinnovo", d.get("Scadenza Rinnovo")),
        ]

        card = ctk.CTkFrame(body, fg_color=THEME["card"], corner_radius=10,
                            border_color=THEME["bordo"], border_width=1)
        card.pack(fill="both", expand=True)
        g = ctk.CTkFrame(card, fg_color="transparent")
        g.pack(fill="both", expand=True, padx=14, pady=10)
        g.columnconfigure(1, weight=1)
        g.columnconfigure(3, weight=1)
        for i, (lbl, val) in enumerate(campi):
            r, cl = i // 2, (i % 2) * 2
            ctk.CTkLabel(g, text=lbl, font=ctk.CTkFont(size=10),
                         text_color=THEME["testo_light"]
                         ).grid(row=r, column=cl, sticky="w", padx=(0,6), pady=3)
            ctk.CTkLabel(g, text=str(val or "—"),
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=THEME["testo"]
                         ).grid(row=r, column=cl+1, sticky="w", pady=3, padx=(0,16))

        btn_row = ctk.CTkFrame(self, fg_color=THEME["sfondo"], corner_radius=0)
        btn_row.pack(fill="x", padx=14, pady=(0,14))
        ctk.CTkButton(btn_row, text="✖  Annulla",
                      command=self.destroy,
                      fg_color=THEME["bordo"], text_color=THEME["testo"],
                      hover_color="#CBD5E1", corner_radius=8, height=36,
                      font=ctk.CTkFont(size=12)).pack(side="left", padx=(0,8))
        ctk.CTkButton(btn_row, text="✅  Conferma e Salva",
                      command=self._conferma,
                      fg_color=THEME["verde"], hover_color=THEME["verde_dark"],
                      corner_radius=8, height=36,
                      font=ctk.CTkFont(size=12, weight="bold")
                      ).pack(side="left", fill="x", expand=True)

    def _conferma(self):
        self.destroy()
        self._on_conferma(self._dati)


# ─────────────────────────────────────────────
#  POPUP ISTRUZIONI BOOKMARKLET
# ─────────────────────────────────────────────
class PopupBookmarklet(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Imposta Bookmarklet")
        self.geometry("620x480")
        self.resizable(False, False)
        self.lift()
        self.focus_force()
        self._build()

    def _build(self):
        hdr = ctk.CTkFrame(self, fg_color=THEME["blu_scuro"],
                           corner_radius=0, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="  📎  Configura il Bookmarklet",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color="white").pack(side="left", padx=14)

        body = ctk.CTkFrame(self, fg_color=THEME["sfondo"], corner_radius=0)
        body.pack(fill="both", expand=True, padx=16, pady=14)

        istruzioni = (
            "Il bookmarklet è un piccolo pulsante che aggiungi ai preferiti del browser.\n"
            "Quando sei sulla pagina di una pratica, cliccalo e i dati\n"
            "vengono inviati automaticamente all'app."
        )
        ctk.CTkLabel(body, text=istruzioni,
                     font=ctk.CTkFont(size=12),
                     text_color=THEME["testo"],
                     justify="left").pack(anchor="w", pady=(0,12))

        # Passi
        passi = [
            ("1", "Copia il codice qui sotto (Ctrl+A poi Ctrl+C)"),
            ("2", "Apri il browser → mostra la barra dei preferiti (Ctrl+Shift+B)"),
            ("3", "Trascina un link qualsiasi nella barra, poi clicca 'Modifica'"),
            ("4", "Cancella l'URL e incolla il codice copiato — salva"),
            ("5", "Ora vai su una pratica e clicca il bookmark!"),
        ]
        for num, testo in passi:
            row = ctk.CTkFrame(body, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(row, text=num,
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color="white",
                         fg_color=THEME["blu_medio"],
                         corner_radius=10,
                         width=22, height=22).pack(side="left", padx=(0,8))
            ctk.CTkLabel(row, text=testo,
                         font=ctk.CTkFont(size=11),
                         text_color=THEME["testo"],
                         anchor="w").pack(side="left")

        ctk.CTkLabel(body, text="Codice del bookmarklet:",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=THEME["blu_scuro"]).pack(anchor="w", pady=(12,4))

        txt = ctk.CTkTextbox(body, height=80,
                             font=ctk.CTkFont(family="Consolas", size=9),
                             fg_color=THEME["card"],
                             border_color=THEME["bordo"],
                             border_width=1, corner_radius=6)
        txt.pack(fill="x", pady=(0,8))
        txt.insert("1.0", BOOKMARKLET_JS)
        txt.configure(state="disabled")

        ctk.CTkButton(body, text="📋  Copia negli appunti",
                      command=lambda: self._copia(txt),
                      fg_color=THEME["blu_medio"], hover_color=THEME["blu_scuro"],
                      corner_radius=8, height=34,
                      font=ctk.CTkFont(size=12, weight="bold")
                      ).pack(fill="x")

    def _copia(self, txt):
        self.clipboard_clear()
        self.clipboard_append(BOOKMARKLET_JS)
        messagebox.showinfo("Copiato",
                            "Codice copiato negli appunti!\n"
                            "Ora incollalo come URL di un preferito nel browser.")


# ─────────────────────────────────────────────
#  WIDGET RIUTILIZZABILI
# ─────────────────────────────────────────────
class Card(ctk.CTkFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=THEME["card"], corner_radius=10,
                         border_color=THEME["bordo"], border_width=1, **kw)

def lbl_f(parent, text):
    return ctk.CTkLabel(parent, text=text, font=ctk.CTkFont(size=10),
                        text_color=THEME["testo_light"])

def campo_v(parent, text):
    val = str(text or "")
    e   = ctk.CTkEntry(parent, font=ctk.CTkFont(size=12),
                       fg_color="transparent", border_width=0,
                       text_color=THEME["testo"] if val else THEME["testo_light"])
    e.insert(0, val if val else "—")
    e.configure(state="readonly")
    return e


# ─────────────────────────────────────────────
#  VISTA LISTA
# ─────────────────────────────────────────────
class VistaLista(ctk.CTkFrame):
    COLONNE = [
        ("Nr. Pratica",    "Nr. Pratica",    110),
        ("Cognome",        "Cognome",         120),
        ("Nome",           "Nome",            120),
        ("Data Nascita",   "Data Nascita",     95),
        ("Cittadinanza",   "Cittadinanza",    110),
        ("CUI",            "CUI",              85),
        ("Stato",          "_stato",          145),
        ("Data Estrazione","Data Estrazione", 130),
    ]

    def __init__(self, parent, get_path, on_select, **kw):
        super().__init__(parent, fg_color=THEME["sfondo"], corner_radius=0, **kw)
        self._get_path  = get_path
        self._on_select = on_select
        self._pratiche  = []
        self._stati     = {}
        self._sort_col  = "Nr. Pratica"
        self._sort_asc  = True
        self._build()

    def _build(self):
        # ── Riga 1: titolo + ricerca ──
        bar1 = ctk.CTkFrame(self, fg_color=THEME["card"],
                            corner_radius=0, height=48)
        bar1.pack(fill="x")
        bar1.pack_propagate(False)

        ctk.CTkLabel(bar1, text="Lista Pratiche",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=THEME["blu_scuro"]
                     ).pack(side="left", padx=14)

        self._search_var = ctk.StringVar()
        self._search_var.trace_add("write", lambda *a: self._filtra())
        ctk.CTkEntry(bar1, textvariable=self._search_var,
                     placeholder_text="🔍  Cerca nome, cognome, CUI, nr...",
                     fg_color=THEME["sfondo"], border_color=THEME["bordo"],
                     height=30, width=260, corner_radius=8,
                     font=ctk.CTkFont(size=11)
                     ).pack(side="left", padx=(0,6))

        ctk.CTkButton(bar1, text="✕",
                      command=lambda: self._search_var.set(""),
                      fg_color=THEME["bordo"], text_color=THEME["testo_light"],
                      hover_color="#CBD5E1", height=30, width=32,
                      corner_radius=8, font=ctk.CTkFont(size=11)
                      ).pack(side="left", padx=(0,8))

        ctk.CTkButton(bar1, text="↻  Aggiorna",
                      command=self.carica,
                      fg_color=THEME["blu_medio"], hover_color=THEME["blu_scuro"],
                      height=30, width=90, corner_radius=8,
                      font=ctk.CTkFont(size=11)
                      ).pack(side="left")

        self.lbl_count = ctk.CTkLabel(bar1, text="",
                                      font=ctk.CTkFont(size=11),
                                      text_color=THEME["testo_light"])
        self.lbl_count.pack(side="right", padx=14)

        # ── Riga 2: filtri ──
        bar2 = ctk.CTkFrame(self, fg_color=THEME["sfondo"],
                            corner_radius=0, height=36)
        bar2.pack(fill="x")
        bar2.pack_propagate(False)

        ctk.CTkLabel(bar2, text="Filtri:",
                     font=ctk.CTkFont(size=11),
                     text_color=THEME["testo_light"]
                     ).pack(side="left", padx=(14,6))

        stati_opzioni = ["Tutti gli stati"] + [STATI[k][0] for k in STATI]
        self._stato_var = ctk.StringVar(value="Tutti gli stati")
        ctk.CTkOptionMenu(bar2, variable=self._stato_var,
                          values=stati_opzioni,
                          fg_color=THEME["card"], button_color=THEME["blu_medio"],
                          text_color=THEME["testo"], font=ctk.CTkFont(size=11),
                          width=160, height=26, corner_radius=8,
                          command=lambda v: self._filtra()
                          ).pack(side="left", padx=(0,8))

        ctk.CTkLabel(bar2, text="Operatore:",
                     font=ctk.CTkFont(size=11),
                     text_color=THEME["testo_light"]
                     ).pack(side="left", padx=(0,4))
        self._op_var  = ctk.StringVar(value="Tutti")
        self._op_menu = ctk.CTkOptionMenu(bar2, variable=self._op_var,
                          values=["Tutti"],
                          fg_color=THEME["card"], button_color=THEME["blu_medio"],
                          text_color=THEME["testo"], font=ctk.CTkFont(size=11),
                          width=130, height=26, corner_radius=8,
                          command=lambda v: self._filtra()
                          )
        self._op_menu.pack(side="left", padx=(0,8))

        ctk.CTkLabel(bar2, text="Scadenza:",
                     font=ctk.CTkFont(size=11),
                     text_color=THEME["testo_light"]
                     ).pack(side="left", padx=(0,4))
        self._scad_var = ctk.StringVar(value="Tutte")
        ctk.CTkOptionMenu(bar2, variable=self._scad_var,
                          values=["Tutte","30 giorni","60 giorni","90 giorni","6 mesi"],
                          fg_color=THEME["card"], button_color=THEME["blu_medio"],
                          text_color=THEME["testo"], font=ctk.CTkFont(size=11),
                          width=110, height=26, corner_radius=8,
                          command=lambda v: self._filtra()
                          ).pack(side="left", padx=(0,8))

        ctk.CTkButton(bar2, text="↺ Reset",
                      command=self._reset_filtri,
                      fg_color="transparent", text_color=THEME["blu_medio"],
                      hover_color=THEME["sfondo"], height=26, corner_radius=8,
                      font=ctk.CTkFont(size=11)
                      ).pack(side="left")

        # ── Tabella ──
        frame = tk.Frame(self, bg=THEME["sfondo"])
        frame.pack(fill="both", expand=True)

        vsb = tk.Scrollbar(frame, orient="vertical")
        vsb.pack(side="right", fill="y")
        hsb = tk.Scrollbar(frame, orient="horizontal")
        hsb.pack(side="bottom", fill="x")

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Lista.Treeview",
                        background=THEME["card"], foreground=THEME["testo"],
                        rowheight=26, fieldbackground=THEME["card"],
                        borderwidth=0, font=("Segoe UI", 10))
        style.configure("Lista.Treeview.Heading",
                        background=THEME["blu_scuro"], foreground="white",
                        relief="flat", font=("Segoe UI", 10, "bold"), padding=(8,6))
        style.map("Lista.Treeview",
                  background=[("selected", THEME["blu_medio"])],
                  foreground=[("selected", "white")])
        style.map("Lista.Treeview.Heading",
                  background=[("active", THEME["blu_medio"])])

        cols = [c[0] for c in self.COLONNE]
        self.tree = ttk.Treeview(frame, columns=cols, show="headings",
                                 style="Lista.Treeview",
                                 yscrollcommand=vsb.set,
                                 xscrollcommand=hsb.set)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        for label, _, width in self.COLONNE:
            self.tree.heading(label, text=label,
                              command=lambda c=label: self._ordina(c))
            self.tree.column(label, width=width, minwidth=50, anchor="w")

        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._apri)
        self.tree.bind("<Return>",   self._apri)

        self.tree.tag_configure("da_verificare", background="#FEF08A")
        self.tree.tag_configure("sospesa",       background="#FED7AA")
        self.tree.tag_configure("validata",      background="#BBF7D0")
        self.tree.tag_configure("negata",        background="#FECACA")
        self.tree.tag_configure("alt",           background="#F1F5F9")

        foot = ctk.CTkFrame(self, fg_color=THEME["card"],
                            corner_radius=0, height=26)
        foot.pack(fill="x")
        foot.pack_propagate(False)
        ctk.CTkLabel(foot,
                     text="  Doppio click o Invio per aprire  •  Click intestazione per ordinare",
                     font=ctk.CTkFont(size=10),
                     text_color=THEME["testo_light"]).pack(side="left")

    def carica(self):
        path = self._get_path()
        self._pratiche = leggi_pratiche(path)
        self._stati    = leggi_tutti_stati(path)
        operatori = sorted(set(
            str(p.get("Estratto Da","") or "").strip()
            for p in self._pratiche if p.get("Estratto Da")
        ))
        self._op_menu.configure(values=["Tutti"] + operatori)
        self._filtra()

    def _reset_filtri(self):
        self._stato_var.set("Tutti gli stati")
        self._op_var.set("Tutti")
        self._scad_var.set("Tutte")
        self._search_var.set("")

    def _filtra(self):
        q     = self._search_var.get().lower()
        esiti = {STATI[k][0]: k for k in STATI}
        fkey  = esiti.get(self._stato_var.get(), "tutte")
        fop   = self._op_var.get()
        fscad = self._scad_var.get()
        oggi  = date.today()
        scad_giorni = {"30 giorni":30,"60 giorni":60,"90 giorni":90,"6 mesi":180}

        out = []
        for p in self._pratiche:
            if q and not any(q in str(p.get(c,"")).lower()
                             for c in ["Cognome","Nome","CUI","Nr. Pratica","Cittadinanza"]):
                continue
            nr    = str(p.get("Nr. Pratica","")).strip()
            stato = self._stati.get(nr, "da_verificare")
            if fkey != "tutte" and stato != fkey:
                continue
            if fop != "Tutti" and str(p.get("Estratto Da","")).strip() != fop:
                continue
            if fscad in scad_giorni:
                try:
                    scad_d = datetime.strptime(str(p.get("Scadenza Rinnovo","")),"%d/%m/%Y").date()
                    if not (oggi <= scad_d <= oggi + timedelta(days=scad_giorni[fscad])):
                        continue
                except Exception:
                    continue
            out.append((p, stato))

        col_map = {c[0]: c[1] for c in self.COLONNE}
        campo   = col_map.get(self._sort_col, self._sort_col)

        def sort_key(item):
            p, stato = item
            if campo == "_stato":
                return STATI.get(stato, STATI["da_verificare"])[0]
            return str(p.get(campo,"") or "").lower()

        out.sort(key=sort_key, reverse=not self._sort_asc)
        self._render(out)

    def _render(self, righe):
        self.tree.delete(*self.tree.get_children())
        self.lbl_count.configure(text=f"{len(righe)} pratiche")
        for i, (p, stato) in enumerate(righe):
            nr      = str(p.get("Nr. Pratica",""))
            info_st = STATI.get(stato, STATI["da_verificare"])
            valori  = (
                nr,
                str(p.get("Cognome","") or ""),
                str(p.get("Nome","") or ""),
                str(p.get("Data Nascita","") or ""),
                str(p.get("Cittadinanza","") or p.get("Nazione Nascita","") or ""),
                str(p.get("CUI","") or ""),
                info_st[0],
                str(p.get("Data Estrazione","") or ""),
            )
            tag = stato if stato in STATI else ("alt" if i % 2 else "")
            self.tree.insert("", "end", iid=nr, values=valori, tags=(tag,))

    def _ordina(self, col):
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        self._filtra()

    def _apri(self, event=None):
        sel = self.tree.selection()
        if sel:
            self._on_select(sel[0])


# ─────────────────────────────────────────────
#  VISTA SCHEDA + ATTIVITÀ
# ─────────────────────────────────────────────
class VistaScheda(ctk.CTkFrame):
    def __init__(self, parent, get_path, vai_lista, **kw):
        super().__init__(parent, fg_color=THEME["sfondo"], corner_radius=0, **kw)
        self._get_path  = get_path
        self._vai_lista = vai_lista
        self._nr        = None
        self._stato_var  = None
        self._stato_menu = None
        self._build()

    def _build(self):
        nav = ctk.CTkFrame(self, fg_color=THEME["card"],
                           corner_radius=0, height=44)
        nav.pack(fill="x")
        nav.pack_propagate(False)

        ctk.CTkButton(nav, text="☰  Lista Pratiche",
                      command=self._vai_lista,
                      fg_color="transparent", text_color=THEME["blu_medio"],
                      hover_color=THEME["sfondo"], height=32, corner_radius=8,
                      font=ctk.CTkFont(size=12)
                      ).pack(side="left", padx=(8,0))

        ctk.CTkLabel(nav, text="›",
                     font=ctk.CTkFont(size=14),
                     text_color=THEME["testo_light"]).pack(side="left", padx=4)

        self.lbl_breadcrumb = ctk.CTkLabel(nav, text="—",
                                           font=ctk.CTkFont(size=12, weight="bold"),
                                           text_color=THEME["testo"])
        self.lbl_breadcrumb.pack(side="left")

        self._corpo = ctk.CTkFrame(self, fg_color=THEME["sfondo"], corner_radius=0)
        self._corpo.pack(fill="both", expand=True)

        self._frame_scheda = ctk.CTkFrame(self._corpo,
                                          fg_color=THEME["sfondo"], corner_radius=0)
        self._frame_scheda.pack(side="left", fill="both", expand=True)

        ctk.CTkFrame(self._corpo, fg_color=THEME["bordo"],
                     width=1, corner_radius=0).pack(side="left", fill="y")

        self._frame_att = ctk.CTkFrame(self._corpo,
                                       fg_color=THEME["sfondo"],
                                       corner_radius=0, width=340)
        self._frame_att.pack(side="left", fill="y")
        self._frame_att.pack_propagate(False)

        ctk.CTkLabel(self._frame_scheda,
                     text="← Seleziona una pratica dalla lista",
                     font=ctk.CTkFont(size=13),
                     text_color=THEME["testo_light"]).pack(expand=True)

    def carica(self, nr: str):
        self._nr = nr
        path     = self._get_path()
        pratiche = leggi_pratiche(path)
        pratica  = next((p for p in pratiche
                         if str(p.get("Nr. Pratica","")).strip() == nr), None)
        if not pratica:
            return
        att = leggi_attivita(path, nr)
        pratica["Stato"] = att.get("Stato","da_verificare") or "da_verificare"
        nome = f"{pratica.get('Cognome','')} {pratica.get('Nome','')}".strip()
        self.lbl_breadcrumb.configure(text=f"{nr}  —  {nome}")
        self._render_scheda(pratica, att)
        self._render_attivita(nr, att)

    def _render_scheda(self, pratica, att):
        for w in self._frame_scheda.winfo_children():
            w.destroy()
        scroll = ctk.CTkScrollableFrame(self._frame_scheda,
                                        fg_color=THEME["sfondo"], corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=14, pady=12)
        nr = str(pratica.get("Nr. Pratica","") or "")

        # Nr. Pratica + Stato
        nr_card = Card(scroll)
        nr_card.pack(fill="x", pady=(0,8))
        ni = ctk.CTkFrame(nr_card, fg_color="transparent")
        ni.pack(fill="x", padx=14, pady=12)
        ctk.CTkLabel(ni, text="Nr. Pratica",
                     font=ctk.CTkFont(size=10),
                     text_color=THEME["testo_light"]).pack(anchor="w")
        ctk.CTkLabel(ni, text=nr,
                     font=ctk.CTkFont(family="Consolas", size=22, weight="bold"),
                     text_color=THEME["blu_scuro"]).pack(anchor="w")

        stato_key  = str(pratica.get("Stato","da_verificare") or "da_verificare")
        info_stato = STATI.get(stato_key, STATI["da_verificare"])
        self._stato_var = ctk.StringVar(value=info_stato[0])
        esiti_map = {STATI[k][0]: k for k in STATI}

        stato_row = ctk.CTkFrame(ni, fg_color="transparent")
        stato_row.pack(anchor="w", fill="x", pady=(8,0))
        self._stato_menu = ctk.CTkOptionMenu(
            stato_row, variable=self._stato_var,
            values=[STATI[k][0] for k in STATI],
            fg_color=info_stato[1], button_color=info_stato[2],
            text_color=info_stato[2],
            font=ctk.CTkFont(size=11, weight="bold"),
            height=30, width=175, corner_radius=16,
            command=lambda v: self._aggiorna_colore(esiti_map.get(v,"da_verificare"))
        )
        self._stato_menu.pack(side="left", padx=(0,10))
        ctk.CTkButton(stato_row, text="💾 Salva stato",
                      command=lambda: self._salva_stato(nr, esiti_map),
                      fg_color=THEME["verde"], hover_color=THEME["verde_dark"],
                      height=30, corner_radius=8, font=ctk.CTkFont(size=11)
                      ).pack(side="left")

        ctk.CTkLabel(nr_card,
                     text=f"Estratto da {pratica.get('Estratto Da','')}  •  {pratica.get('Data Estrazione','')}",
                     font=ctk.CTkFont(size=9),
                     text_color=THEME["testo_light"]
                     ).pack(anchor="e", padx=14, pady=(0,8))

        # Anagrafica
        ana = Card(scroll)
        ana.pack(fill="x", pady=(0,8))
        ctk.CTkLabel(ana, text="👤  Anagrafica",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=THEME["blu_scuro"]
                     ).pack(anchor="w", padx=14, pady=(10,4))
        ctk.CTkFrame(ana, fg_color=THEME["bordo"], height=1, corner_radius=0
                     ).pack(fill="x", padx=14, pady=(0,8))
        campi = [
            ("Cognome",          pratica.get("Cognome")),
            ("Nome",             pratica.get("Nome")),
            ("CUI",              pratica.get("CUI")),
            ("Codice Fiscale",   pratica.get("Codice Fiscale")),
            ("Data Nascita",     pratica.get("Data Nascita")),
            ("Luogo Nascita",    pratica.get("Luogo Nascita")),
            ("Nazione Nascita",  pratica.get("Nazione Nascita")),
            ("Cittadinanza",     pratica.get("Cittadinanza")),
            ("Sesso",            pratica.get("Sesso")),
            ("Stato Civile",     pratica.get("Stato Civile")),
            ("Telefono",         pratica.get("Telefono")),
            ("Coniuge",          pratica.get("Coniuge")),
            ("Comune",           pratica.get("Comune")),
            ("Indirizzo",        pratica.get("Indirizzo")),
            ("Tipo Soggiorno",   pratica.get("Tipo Soggiorno")),
            ("Tipo Pratica",     pratica.get("Tipo Pratica")),
            ("Motivo Soggiorno", pratica.get("Motivo Soggiorno")),
            ("Validita Sogg.",   pratica.get("Validita Soggiorno")),
            ("Presentazione",    pratica.get("Data Presentazione")),
            ("Scadenza Rinnovo", pratica.get("Scadenza Rinnovo")),
            ("Referenze",        pratica.get("Referenze")),
            ("Op. Gestionale",   pratica.get("Operatore Gestionale")),
        ]
        g = ctk.CTkFrame(ana, fg_color="transparent")
        g.pack(fill="x", padx=14, pady=(0,12))
        g.columnconfigure(1, weight=1)
        g.columnconfigure(3, weight=1)
        for i, (lbl, val) in enumerate(campi):
            r, cl = i // 2, (i % 2) * 2
            lbl_f(g, lbl).grid(row=r, column=cl, sticky="nw", padx=(0,6), pady=2)
            campo_v(g, val).grid(row=r, column=cl+1, sticky="ew", pady=2, padx=(0,16))

        if pratica.get("Note Gestionale"):
            nc = Card(scroll)
            nc.pack(fill="x", pady=(0,8))
            ctk.CTkLabel(nc, text="📝  Note gestionale",
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=THEME["blu_scuro"]
                         ).pack(anchor="w", padx=14, pady=(10,4))
            ctk.CTkFrame(nc, fg_color=THEME["bordo"], height=1, corner_radius=0
                         ).pack(fill="x", padx=14, pady=(0,6))
            ctk.CTkLabel(nc, text=str(pratica.get("Note Gestionale","")),
                         font=ctk.CTkFont(size=11), text_color=THEME["testo"],
                         wraplength=380, justify="left"
                         ).pack(anchor="w", padx=14, pady=(0,12))

    def _render_attivita(self, nr: str, att: dict):
        for w in self._frame_att.winfo_children():
            w.destroy()
        hdr = ctk.CTkFrame(self._frame_att, fg_color=THEME["blu_scuro"],
                           corner_radius=0, height=44)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="Attività",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color="white").pack(side="left", padx=14, pady=10)

        tab = ctk.CTkTabview(self._frame_att, fg_color=THEME["sfondo"],
                             segmented_button_fg_color=THEME["card"],
                             segmented_button_selected_color=THEME["blu_medio"],
                             segmented_button_selected_hover_color=THEME["blu_scuro"],
                             segmented_button_unselected_color=THEME["card"],
                             text_color=THEME["testo"], corner_radius=0)
        tab.pack(fill="both", expand=True)
        tab.add("📝 Note")
        tab.add("✅ Checklist")
        self._build_note(tab.tab("📝 Note"), nr, att)
        self._build_checklist(tab.tab("✅ Checklist"), nr, att)

    def _build_note(self, parent, nr, att):
        scroll = ctk.CTkScrollableFrame(parent, fg_color=THEME["sfondo"], corner_radius=0)
        scroll.pack(fill="both", expand=True)
        nc = Card(scroll)
        nc.pack(fill="x", padx=6, pady=(6,4))
        ctk.CTkLabel(nc, text="📝  La mia nota",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=THEME["blu_scuro"]
                     ).pack(anchor="w", padx=10, pady=(8,4))
        self._nota_txt = ctk.CTkTextbox(nc, height=240,
                                        font=ctk.CTkFont(size=12),
                                        fg_color=THEME["sfondo"],
                                        border_color=THEME["bordo"],
                                        border_width=1, corner_radius=6)
        self._nota_txt.pack(fill="x", padx=10, pady=(0,8))
        self._nota_txt.insert("1.0", str(att.get("Note Personali","") or ""))
        ctk.CTkButton(nc, text="💾  Salva nota",
                      command=lambda: self._salva_nota(nr),
                      fg_color=THEME["blu_medio"], hover_color=THEME["blu_scuro"],
                      corner_radius=8, height=32, font=ctk.CTkFont(size=11)
                      ).pack(anchor="e", padx=10, pady=(0,10))
        if att.get("Note Modificate Il"):
            ctk.CTkLabel(scroll,
                         text=f"Ultima modifica: {att['Note Modificate Il']}",
                         font=ctk.CTkFont(size=9),
                         text_color=THEME["testo_light"]).pack(pady=2)

    def _build_checklist(self, parent, nr, att):
        scroll = ctk.CTkScrollableFrame(parent, fg_color=THEME["sfondo"], corner_radius=0)
        scroll.pack(fill="both", expand=True)

        def sezione(titolo):
            c = Card(scroll)
            c.pack(fill="x", padx=6, pady=(6,3))
            ctk.CTkLabel(c, text=titolo,
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=THEME["blu_scuro"]
                         ).pack(anchor="w", padx=10, pady=(8,4))
            return c

        s1 = sezione("SDI")
        lbl_f(s1,"Esito").pack(anchor="w", padx=10)
        self._sdi_esito = ctk.CTkOptionMenu(s1, values=SDI_ESITI,
                                            fg_color=THEME["sfondo"],
                                            button_color=THEME["blu_medio"],
                                            text_color=THEME["testo"],
                                            font=ctk.CTkFont(size=11),
                                            height=28, corner_radius=6)
        self._sdi_esito.set(str(att.get("SDI Esito","") or ""))
        self._sdi_esito.pack(fill="x", padx=10, pady=(2,6))
        lbl_f(s1,"Note SDI").pack(anchor="w", padx=10)
        self._sdi_note = ctk.CTkTextbox(s1, height=52,
                                        font=ctk.CTkFont(size=11),
                                        fg_color=THEME["sfondo"],
                                        border_color=THEME["bordo"],
                                        border_width=1, corner_radius=6)
        self._sdi_note.pack(fill="x", padx=10, pady=(2,10))
        self._sdi_note.insert("1.0", str(att.get("SDI Note","") or ""))

        s2 = sezione("Reddito")
        lbl_f(s2,"Tipo").pack(anchor="w", padx=10)
        self._reddito_tipo = ctk.CTkOptionMenu(s2, values=REDDITO_TIPI,
                                               fg_color=THEME["sfondo"],
                                               button_color=THEME["blu_medio"],
                                               text_color=THEME["testo"],
                                               font=ctk.CTkFont(size=11),
                                               height=28, corner_radius=6)
        self._reddito_tipo.set(str(att.get("Reddito Tipo","") or ""))
        self._reddito_tipo.pack(fill="x", padx=10, pady=(2,6))
        lbl_f(s2,"Note reddito").pack(anchor="w", padx=10)
        self._reddito_note = ctk.CTkTextbox(s2, height=52,
                                            font=ctk.CTkFont(size=11),
                                            fg_color=THEME["sfondo"],
                                            border_color=THEME["bordo"],
                                            border_width=1, corner_radius=6)
        self._reddito_note.pack(fill="x", padx=10, pady=(2,10))
        self._reddito_note.insert("1.0", str(att.get("Reddito Note","") or ""))

        s3 = sezione("Famiglia / Trainante")
        self._fam_var = ctk.BooleanVar(
            value=str(att.get("Famiglia Trainante","")).upper()=="SI")
        ctk.CTkCheckBox(s3, text="Presenza trainante",
                        variable=self._fam_var,
                        fg_color=THEME["blu_medio"],
                        hover_color=THEME["blu_scuro"],
                        font=ctk.CTkFont(size=11),
                        command=self._toggle_fam
                        ).pack(anchor="w", padx=10, pady=(0,6))
        self._fam_frame = ctk.CTkFrame(s3, fg_color="transparent")
        self._fam_frame.pack(fill="x", padx=10, pady=(0,10))
        self._fam_fields = {}
        for label, key in [("Nome trainante","Nome Trainante"),
                            ("CF trainante","CF Trainante"),
                            ("Nr. pratica trainante","Nr Pratica Trainante")]:
            lbl_f(self._fam_frame, label).pack(anchor="w")
            e = ctk.CTkEntry(self._fam_frame,
                             fg_color=THEME["sfondo"], border_color=THEME["bordo"],
                             height=28, corner_radius=6, font=ctk.CTkFont(size=11))
            e.pack(fill="x", pady=(2,4))
            e.insert(0, str(att.get(key,"") or ""))
            self._fam_fields[key] = e
        self._toggle_fam()

        s4 = sezione("Documenti mancanti")
        self._doc = ctk.CTkTextbox(s4, height=68,
                                   font=ctk.CTkFont(size=11),
                                   fg_color=THEME["sfondo"],
                                   border_color=THEME["bordo"],
                                   border_width=1, corner_radius=6)
        self._doc.pack(fill="x", padx=10, pady=(0,10))
        self._doc.insert("1.0", str(att.get("Documenti Mancanti","") or ""))

        ctk.CTkButton(scroll, text="💾  Salva Checklist",
                      command=lambda: self._salva_checklist(nr),
                      fg_color=THEME["verde"], hover_color=THEME["verde_dark"],
                      corner_radius=8, height=34,
                      font=ctk.CTkFont(size=12, weight="bold")
                      ).pack(fill="x", padx=6, pady=8)

        if att.get("Checklist Compilata Da"):
            ctk.CTkLabel(scroll,
                         text=f"Compilata da {att['Checklist Compilata Da']}  •  {att.get('Checklist Compilata Il','')}",
                         font=ctk.CTkFont(size=9),
                         text_color=THEME["testo_light"]).pack(pady=(0,4))

    def _toggle_fam(self):
        st = "normal" if self._fam_var.get() else "disabled"
        for e in self._fam_fields.values():
            e.configure(state=st,
                        fg_color=THEME["sfondo"] if st=="normal" else THEME["bordo"])

    def _aggiorna_colore(self, stato_key: str):
        info = STATI.get(stato_key, STATI["da_verificare"])
        self._stato_menu.configure(
            fg_color=info[1], button_color=info[2], text_color=info[2])

    def _salva_stato(self, nr: str, esiti_map: dict):
        nuovo = esiti_map.get(self._stato_var.get(), "da_verificare")
        att   = leggi_attivita(self._get_path(), nr)
        att["Nr. Pratica"] = nr
        att["Stato"]       = nuovo
        if not att.get("SDI Esito"):
            att["SDI Esito"] = ""
        salva_attivita_excel(self._get_path(), att)
        messagebox.showinfo("Salvato", f"Stato aggiornato: {STATI[nuovo][0]}")

    def _salva_nota(self, nr: str):
        att = leggi_attivita(self._get_path(), nr)
        att.update({
            "Nr. Pratica":       nr,
            "Note Personali":    self._nota_txt.get("1.0","end").strip(),
            "Note Modificate Il":datetime.now().strftime("%d/%m/%Y %H:%M"),
        })
        if not att.get("Stato"):
            att["Stato"] = "da_verificare"
        salva_attivita_excel(self._get_path(), att)
        messagebox.showinfo("Salvato","Nota salvata.")

    def _salva_checklist(self, nr: str):
        att = leggi_attivita(self._get_path(), nr)
        att.update({
            "Nr. Pratica":           nr,
            "SDI Esito":             self._sdi_esito.get(),
            "SDI Note":              self._sdi_note.get("1.0","end").strip(),
            "Reddito Tipo":          self._reddito_tipo.get(),
            "Reddito Note":          self._reddito_note.get("1.0","end").strip(),
            "Famiglia Trainante":    "SI" if self._fam_var.get() else "NO",
            "Nome Trainante":        self._fam_fields["Nome Trainante"].get(),
            "CF Trainante":          self._fam_fields["CF Trainante"].get(),
            "Nr Pratica Trainante":  self._fam_fields["Nr Pratica Trainante"].get(),
            "Documenti Mancanti":    self._doc.get("1.0","end").strip(),
            "Checklist Compilata Da":OPERATORE,
            "Checklist Compilata Il":datetime.now().strftime("%d/%m/%Y %H:%M"),
        })
        if not att.get("Stato"):
            att["Stato"] = "da_verificare"
        salva_attivita_excel(self._get_path(), att)
        messagebox.showinfo("Salvato","Checklist salvata.")


# ─────────────────────────────────────────────
#  APP PRINCIPALE
# ─────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Pratiche Immigrazione  —  {OPERATORE}")
        self.geometry("1200x740")
        self.minsize(900, 580)
        self._excel_path = ctk.StringVar(value=EXCEL_PATH_DEFAULT)
        self._server     = None
        self._build()
        self._avvia_server()
        self._mostra_lista()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build(self):
        # Topbar
        top = ctk.CTkFrame(self, fg_color=THEME["blu_scuro"],
                           corner_radius=0, height=50)
        top.pack(fill="x")
        top.pack_propagate(False)

        ctk.CTkLabel(top, text="  🏛  Pratiche Immigrazione",
                     font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
                     text_color="white").pack(side="left", padx=10)

        # File Excel
        ctk.CTkEntry(top, textvariable=self._excel_path,
                     fg_color="#1a3f63", border_color=THEME["blu_medio"],
                     text_color="white", height=30, width=240, corner_radius=6,
                     font=ctk.CTkFont(size=10)
                     ).pack(side="left", padx=(16,4))
        ctk.CTkButton(top, text="📁", width=32, height=30,
                      fg_color=THEME["blu_medio"], hover_color="#1a3f63",
                      corner_radius=6, command=self._scegli_file
                      ).pack(side="left", padx=(0,4))

        # Bottone bookmarklet
        ctk.CTkButton(top, text="📎  Bookmarklet",
                      command=self._mostra_bookmarklet,
                      fg_color=THEME["arancio"], hover_color="#C2410C",
                      height=30, corner_radius=8,
                      font=ctk.CTkFont(size=11)
                      ).pack(side="left", padx=(12,0))

        # Indicatore server
        self.lbl_server = ctk.CTkLabel(top, text="⚪ Server...",
                                       font=ctk.CTkFont(size=10),
                                       text_color="#93C5FD")
        self.lbl_server.pack(side="left", padx=10)

        ctk.CTkLabel(top, text=f"👤  {OPERATORE}",
                     font=ctk.CTkFont(size=11),
                     text_color="#93C5FD").pack(side="right", padx=14)

        # Contenitore viste
        self._container = ctk.CTkFrame(self, fg_color=THEME["sfondo"], corner_radius=0)
        self._container.pack(fill="both", expand=True)

        self._vista_lista = VistaLista(
            self._container,
            get_path=self._excel_path.get,
            on_select=self._apri_pratica
        )
        self._vista_scheda = VistaScheda(
            self._container,
            get_path=self._excel_path.get,
            vai_lista=self._mostra_lista
        )

    def _avvia_server(self):
        try:
            self._server = ServerLocale(callback=self._ricezione_dati)
            self._server.avvia()
            self.lbl_server.configure(text=f"🟢 In ascolto :{SERVER_PORT}")
        except OSError:
            self.lbl_server.configure(text="🔴 Porta occupata")

    def _ricezione_dati(self, dati: dict):
        """Chiamato dal server quando arrivano dati dal bookmarklet."""
        dati["Estratto Da"]    = OPERATORE
        dati["Data Estrazione"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        self.after(0, self._mostra_anteprima, dati)

    def _mostra_anteprima(self, dati: dict):
        self.lift()
        self.focus_force()
        PopupAnteprima(self, dati, on_conferma=self._salva_e_apri)

    def _salva_e_apri(self, dati: dict):
        try:
            salva_pratica_excel(self._excel_path.get(), dati)
            nr = dati.get("Nr. Pratica","")
            self._mostra_lista()
            self.after(200, lambda: self._apri_pratica(nr))
        except Exception as e:
            messagebox.showerror("Errore salvataggio", str(e))

    def _mostra_bookmarklet(self):
        PopupBookmarklet(self)

    def _mostra_lista(self):
        self._vista_scheda.pack_forget()
        self._vista_lista.pack(fill="both", expand=True)
        self._vista_lista.carica()

    def _apri_pratica(self, nr: str):
        self._vista_lista.pack_forget()
        self._vista_scheda.pack(fill="both", expand=True)
        self._vista_scheda.carica(nr)

    def _scegli_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
        if p:
            self._excel_path.set(p)
            self._mostra_lista()

    def _on_close(self):
        if self._server:
            self._server.ferma()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
