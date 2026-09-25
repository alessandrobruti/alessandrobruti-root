# -*- coding: utf-8 -*-
"""Genera KPI_Tree.md, Data_Request_LAica.xlsx, Agenda_Data_Discovery.md da kpi_data.py"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from kpi_data import KPI, OWN, STAGE, MOTION, wbs_phase
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.join(os.path.dirname(__file__), "..")
PRIO = {"M": "Must", "S": "Should", "C": "Could"}
LVL = {"L1": "Business", "L2": "Servizio", "L3": "Operativo"}

# ---------- KPI_Tree.md ----------
md = ["# Albero KPI Laica: modello Bow Tie × motion\n",
"Aggiornato al 25/09/2026. Fonte unica: `build/kpi_data.py`, che genera anche l'Excel e l'agenda. I rimandi `C-xx` puntano a `Registro_Affermazioni.md`.\n",
"## Logica\n",
"```",
"L1 BUSINESS   ricavi · margine per canale · crescita · retention/NRR · LTV · CAC · payback ≤24m",
"   ▲ spiegato da",
"L2 SERVIZIO   OTIF/fill rate · resi · tempi CS · motivi di contatto · rating · sell-out",
"   ▲ spiegato da",
"L3 OPERATIVO  traffico · conversioni per stadio · ACoS/TACoS · pipeline OEM · ore manuali · data quality · time-to-market",
"```\n",
"Stadi Bow Tie: 1 Awareness → 2 Education → 3 Selection → 4 **Commit** → 5 Onboarding → 6 Impact → 7 Growth.",
"Quattro motion con logiche diverse, da non mescolare:",
"- **RET**: B2B2C retail e distributori (sell-in);",
"- **OEM**: white label, ciclo 1-3 anni;",
"- **AMZ**: Amazon Seller/Vendor, cliente non nostro;",
"- **D2C**: shop proprio.\n",
"## Mappa copertura (n. KPI per motion × stadio)\n"]
stages = ["1", "2", "3", "4", "5", "6", "7", "X"]
def st_hits(s, st):
    if st == "X": return s == "X"
    if "-" in s:
        a, b = s.split("-"); return int(a) <= int(st) <= int(b)
    return s == st
md.append("| Motion | " + " | ".join(STAGE[s] for s in stages) + " |")
md.append("|---|" + "---|" * len(stages))
for m in ["RET", "OEM", "AMZ", "D2C", "ALL"]:
    row = []
    for st in stages:
        n = sum(1 for k in KPI if k[1] == m and st_hits(k[2], st))
        row.append(str(n) if n else "·")
    md.append(f"| {MOTION[m]} | " + " | ".join(row) + " |")
md.append("\nL'Onboarding (5) è scoperto quasi ovunque: oggi Laica non sa chi ha comprato né se il prodotto è stato attivato. È il gap strutturale del Bow Tie destro (vedi K38).\n")

for lvl in ["L1", "L2", "L3"]:
    md.append(f"## {lvl}: KPI {LVL[lvl]}\n")
    md.append("| ID | Motion | Stadio | KPI | Formula | Perché per Laica |")
    md.append("|---|---|---|---|---|---|")
    for k in KPI:
        if k[3] != lvl: continue
        md.append(f"| {k[0]} | {k[1]} | {STAGE[k[2]]} | **{k[4]}** | {k[5]} | {k[6]} |")
    md.append("")

md.append("## Albero causale: i tre KPI che il CFO guarderà\n")
md.append("""```
Margine contributo per canale (K02)
├── Costo canale Amazon % (K33) ← fee 18% · FBA 20-22% · ADS 595k€ (K34) · resi 4,13% (K35)
├── Mix canale (K01) ← quanto Seller migra su D2C (K43) ← registrazioni prodotto (K38) ← packaging/garanzia
└── Costo servizio (K06) ← ticket/1000 unità (K60) × tempo (K61) × % automatizzabile (K62)

LTV consumatore (K45)
├── Tasso di riacquisto filtri a 12 mesi (K45) ← promemoria/abbonamento (K46) ← consenso marketing (K50)
└── Perdita verso compatibili Brita/Philips (C-49) ← share of search (K41)

Crescita OEM (K22 × K25)
├── Prospect ICP raggiunti (K20) → opportunità (K21) → win rate (K22) → durata ciclo (K23)
└── Retention ed espansione OEM (K25)
```
""")
open(os.path.join(ROOT, "KPI_Tree.md"), "w").write("\n".join(md))

# ---------- Excel ----------
wb = Workbook()
thin = Side(style="thin", color="DCDCDC")
HDR = PatternFill("solid", fgColor="000000"); YEL = PatternFill("solid", fgColor="FDC300"); ALT = PatternFill("solid", fgColor="F4F4F4")
F = "Manrope"
def sheet(ws, headers, rows, widths):
    ws.append(headers)
    for c in ws[1]:
        c.fill = HDR; c.font = Font(name=F, bold=True, color="FFFFFF", size=10); c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, r in enumerate(rows):
        ws.append(r)
        for c in ws[ws.max_row]:
            c.font = Font(name=F, size=9); c.alignment = Alignment(wrap_text=True, vertical="top"); c.border = Border(bottom=thin)
            if i % 2: c.fill = ALT
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C2"; ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 30

# Leggimi
ws = wb.active; ws.title = "Leggimi"
info = [["Laica S.p.A.: Richiesta dati RevOps assessment (fase WBS 1.1)"],
        ["Versione", "v1.0 · 25/09/2026 · Impresoft Engage"],
        ["Scadenza consegna", "09/10/2026 (analisi il 15/10, workshop dalla settimana del 19/10)"],
        ["Come usarlo", "Un foglio per referente. Colonna 'Tipo': DIRETTA = sappiamo dove sta il dato, serve l'export; DOMANDA = prima dobbiamo capire se il dato esiste."],
        ["Priorità", "Must = senza questo non si fa l'analisi del 15/10 · Should = serve per i workshop · Could = utile, non bloccante"],
        ["Formato", "CSV o XLSX, una riga per record; periodo 01/2023→oggi dove non indicato; va bene l'estrazione grezza, la puliamo noi."],
        ["Privacy", "Per contatti ed email consumatori va bene l'email pseudonimizzata (hash), purché stabile tra i file. NDA firmato il 05/08/2026."],
        ["Stato", "Colonne 'Stato', 'Consegnato il' e 'Note cliente' da compilare."],
        [], ["Referente", "Nome", "N. richieste", "di cui Must"]]
for r in info: ws.append(r)
ws["A1"].font = Font(name=F, bold=True, size=14); ws.column_dimensions["A"].width = 22; ws.column_dimensions["B"].width = 95
for code, name in OWN.items():
    rows = [k for k in KPI if k[10] == code]
    if rows: ws.append([code, name, len(rows), sum(1 for k in rows if k[13] == "M")])
for row in ws.iter_rows(min_row=2):
    for c in row:
        if c.font.bold is not True: c.font = Font(name=F, size=10); c.alignment = Alignment(wrap_text=True, vertical="top")
for c in ws[10]: c.fill = YEL; c.font = Font(name=F, bold=True)

H = ["ID", "Priorità", "Tipo", "Richiesta / domanda (pronta da inviare)", "Sistema sorgente", "Granularità · periodo", "Dato necessario", "KPI abilitato", "Motion", "Stadio Bow Tie", "Livello", "Fase WBS", "Stato", "Consegnato il", "Note cliente"]
Wd = [6, 8, 10, 70, 22, 20, 36, 34, 14, 16, 10, 26, 10, 12, 30]
def rowof(k):
    return [k[0], PRIO[k[13]], "DIRETTA" if k[11] == "DIR" else "DOMANDA", k[12], k[9], k[8], k[7], k[4], MOTION[k[1]], STAGE[k[2]], LVL[k[3]], wbs_phase(k), "", "", ""]
order = {"M": 0, "S": 1, "C": 2}
allrows = sorted(KPI, key=lambda k: (order[k[13]], k[0]))
sheet(wb.create_sheet("Tutte"), ["Referente"] + H, [[OWN[k[10]]] + rowof(k) for k in allrows], [24] + Wd)
for code, name in OWN.items():
    rows = [rowof(k) for k in allrows if k[10] == code]
    if rows: sheet(wb.create_sheet(code), H, rows, Wd)
# KPI dictionary
sheet(wb.create_sheet("KPI_Dictionary"), ["ID", "Livello", "Motion", "Stadio", "KPI", "Formula", "Perché (rif. registro)", "Owner"],
      [[k[0], LVL[k[3]], MOTION[k[1]], STAGE[k[2]], k[4], k[5], k[6], OWN[k[10]]] for k in KPI], [6, 10, 18, 18, 40, 50, 50, 28])
wb.save(os.path.join(ROOT, "Data_Request_LAica.xlsx"))

# ---------- Agenda ----------
ag = ["# Agenda di data discovery: kick-off del 02/10\n",
"Uso: 1) prima del 02/10 si invia l'Excel con le richieste DIRETTE; 2) al kick-off si fanno le DOMANDE, raggruppate per persona; 3) ogni risposta aggiorna `Stato` nell'Excel e il livello di confidenza nel registro.\n",
"## Da chiarire subito (conflitti tra fonti)\n",
"| # | Domanda | A chi | Perché |",
"|---|---|---|---|",
"| 1 | Nel 95/5 l'Amazon Vendor sta nel B2B o nell'online? Quanto vale in € ciascun canale (retail, OEM, Vendor, Seller, shop)? | Zanuso | Due definizioni diverse (C-01) |",
"| 2 | Il target 15,3M GBP (7,6 + 8,0) a cosa si riferisce: divisione, Laica, incremento, D2C? | Zanuso | I conti non tornano (C-02) |",
"| 3 | Vendita D2C di hardware: quando? (SPICED: FY29/30; Dolcetta: imminente; Milich: urgente) | Dolcetta + Milich insieme | C-27 |",
"| 4 | ERP↔Seller: integrazione live o script CSV? Chi la manutiene? | Sella | C-23 |",
"| 5 | Lo shop fattura \"poche centinaia di k€\" o circa 16 ordini/mese? | Bellero | C-28 |",
"| 6 | Budget AMS di 595 k€/anno: diviso Vendor/Seller, gestito da chi, con quale ACoS/TACoS? | Contin | C-13 |",
"| 7 | Il 18% di fee Amazon come è calcolato (categorie, servizi digitali)? | Zanuso | Tariffa pubblica 15% (C-50) |",
"| 8 | Condividete gli allegati della mail del 02/09 (questionario compilato + ZIP) nella cartella Drive \"00. File condivisi Cliente\"? | Campione (interno) | Fonte primaria non accessibile |",
"| 9 | Condivisione delle 4 registrazioni Bluedot delle call con il cliente | Campione (interno) | Nessun accesso con l'account Bruti |",
""]
for code, name in OWN.items():
    qs = [k for k in allrows if k[10] == code and k[11] == "DISC"]
    if not qs: continue
    ag.append(f"## {name}\n")
    for k in qs:
        ag.append(f"- [ ] **{k[0]}** ({PRIO[k[13]]}): {k[12]}")
    ag.append("")
open(os.path.join(ROOT, "Agenda_Data_Discovery.md"), "w").write("\n".join(ag))
print("ok", len(KPI), sum(1 for k in KPI if k[11] == "DIR"), sum(1 for k in KPI if k[13] == "M"))
