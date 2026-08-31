# -*- coding: utf-8 -*-
"""
Costruisce il file WBS del progetto di Assessment Go-to-Market.
Tutti i calcoli sono formule Excel vive: il file si ricalcola all'apertura.

Identità visiva: Impresoft Brand Manual & Corporate Guidelines (agg. 27.04.2023),
ricetta Excel di references/office-docs.md. Palette di tre colori — giallo
istituzionale, nero, bianco — più grigi puri derivati dal nero. Nessun colore
semantico: gli stati sono codificati con forma, peso e scala monocroma.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment

OUT = "/home/user/alessandrobruti-root/wbs-assessment/WBS_Assessment_GTM.xlsx"

# ---------------------------------------------------------------- brand
# Font istituzionale. Se Manrope non è installato sulle macchine che aprono
# il file, sostituire con "Arial" (fallback indicato dal manuale).
FONT = "Manrope"

YEL  = "FDC300"   # giallo istituzionale
BLK  = "000000"   # nero
WHT  = "FFFFFF"   # bianco
S2   = "F4F4F4"   # surface-2
INK2 = "3A3A3A"   # testo secondario
INK3 = "6E6E6E"   # testo terziario
RULE = "DCDCDC"   # righe sottili
RULS = "9B9B9B"   # righe marcate
G_MID = "8C8C8C"  # grigio della scala grafici (serie non in evidenza)

F_TITLE = Font(name=FONT, size=13, bold=True, color=WHT)
F_SUB   = Font(name=FONT, size=8.5, color=INK3)
F_HEAD  = Font(name=FONT, size=9, bold=True, color=WHT)
F_HEADY = Font(name=FONT, size=9, bold=True, color=BLK)     # header colonna da compilare
F_BODY  = Font(name=FONT, size=9, color=BLK)
F_BODYB = Font(name=FONT, size=9, bold=True, color=BLK)
F_IN    = Font(name=FONT, size=9, color=BLK)                # cella da compilare
F_CALC  = Font(name=FONT, size=9, color=INK2)               # cella calcolata
F_CALCB = Font(name=FONT, size=9, bold=True, color=INK2)
F_TOTAL = Font(name=FONT, size=9, bold=True, color=WHT)
F_EX    = Font(name=FONT, size=9, italic=True, color=INK3)
F_ACC   = Font(name=FONT, size=9, bold=True, color=BLK)     # testo su fondo giallo

# I fill vanno dichiarati in ARGB a 8 cifre con entrambi i colori: openpyxl
# omette l'attributo quando il valore coincide con il proprio default, e
# "000000" normalizzato È quel default — un fill nero dichiarato così
# sparisce dai dxf della formattazione condizionale.
def fill(hex6):
    argb = "FF" + hex6
    return PatternFill("solid", start_color=argb, end_color=argb)

FILL_BLK = fill(BLK)
FILL_YEL = fill(YEL)
FILL_S2  = fill(S2)
FILL_WHT = fill(WHT)
FILL_GMID = fill(G_MID)

# Tabelle: righe orizzontali sottili, nessun bordo verticale.
# Le celle da compilare portano un riquadro nero: è il marcatore di "qui si scrive".
r_thin  = Side(style="thin", color=RULE)
r_black = Side(style="thin", color=BLK)
r_yel   = Side(style="thick", color=YEL)

HRULE   = Border(bottom=r_thin)
INBOX   = Border(left=r_black, right=r_black, top=r_black, bottom=r_black)
r_grey  = Side(style="thin", color=RULS)
CALCBOX = Border(left=r_grey, right=r_grey, top=r_grey, bottom=r_grey)
UNDERY  = Border(bottom=r_yel)          # filetto giallo sotto la fascia titolo
OVERY   = Border(top=r_yel)             # filetto giallo sopra la riga totale

TOP  = Alignment(horizontal="left",   vertical="top", wrap_text=True)
CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)
CTRV = Alignment(horizontal="center", vertical="top")

EUR = '#,##0.00\\ "€";-#,##0.00\\ "€";"-"'
NUM = '#,##0.0;-#,##0.0;"-"'
INT = '#,##0;-#,##0;"-"'

wb = Workbook()


def banda_titolo(sh, rng, testo):
    """Fascia titolo: fondo nero, testo bianco, filetto giallo sotto.
    Il chevron '›' è il pittogramma del marchio."""
    first = rng.split(":")[0]
    sh.merge_cells(rng)
    sh[first] = f"›  {testo}"
    sh[first].font = F_TITLE
    sh[first].alignment = Alignment(horizontal="left", vertical="center")
    col_a, col_b = first[0], rng.split(":")[1][0]
    row = int("".join(ch for ch in first if ch.isdigit()))
    for i in range(ord(col_a), ord(col_b) + 1):
        c = sh[f"{chr(i)}{row}"]
        c.fill = FILL_BLK
        c.border = UNDERY
    sh.row_dimensions[row].height = 26


# ================================================================ PARAMETRI
ps = wb.create_sheet("Parametri")
ps.sheet_view.showGridLines = False

banda_titolo(ps, "A1:E1", "PARAMETRI DI CALCOLO")

ps["A2"] = ("Le celle con riquadro nero, sotto l'intestazione gialla, sono da compilare. Tutto il resto del file si "
            "ricalcola da qui: modificando un valore in questa pagina si aggiornano ore e costi di "
            "tutte le fasi.")
ps["A2"].font = F_SUB
ps.merge_cells("A2:E2")
ps.row_dimensions[2].height = 24

for col, txt, w in (("A", "Parametro", 42), ("B", "", 2), ("C", "Valore", 14),
                    ("D", "Unità", 12), ("E", "Note", 56)):
    ps.column_dimensions[col].width = w
    ps[f"{col}4"].fill = FILL_YEL if col == "C" else FILL_BLK
    if txt:
        ps[f"{col}4"] = txt
        ps[f"{col}4"].font = F_HEADY if col == "C" else F_HEAD
        ps[f"{col}4"].alignment = CTR

PARAMS = [
    (5,  "Ore per giornata", 8, "h",
     "Base di conversione giornate → ore. Confermato: 8 ore.", False),
    (6,  "Tariffa oraria — Front office", None, "€/h",
     "Ore erogate presso il cliente.", True),
    (7,  "Tariffa oraria — Back office", None, "€/h",
     "Ore di lavoro interno (analisi, delivery, reportistica).", True),
    (8,  "Tariffa oraria — Viaggio", None, "€/h",
     "Ore di trasferimento. Incluse nel monte ore complessivo di progetto.", True),
    (10, "Lead time target di progetto", 10, "settimane",
     "Durata di calendario dichiarata in fase di impostazione.", False),
]
for r, label, val, unit, note, _ in PARAMS:
    ps[f"A{r}"] = label;  ps[f"A{r}"].font = F_BODYB; ps[f"A{r}"].alignment = TOP
    c = ps[f"C{r}"]
    c.value = val
    c.font = F_IN; c.fill = FILL_WHT; c.border = INBOX
    c.alignment = CTRV
    c.number_format = INT if unit != "€/h" else EUR
    ps[f"D{r}"] = unit; ps[f"D{r}"].font = F_BODY; ps[f"D{r}"].alignment = CTRV
    ps[f"E{r}"] = note; ps[f"E{r}"].font = F_BODY; ps[f"E{r}"].alignment = TOP
    for col in "ABDE":
        ps[f"{col}{r}"].border = HRULE

ps["C6"].comment = Comment("Tariffa non ancora fornita. Finché la cella resta vuota, "
                           "la colonna «Costo totale» della WBS vale 0.", "WBS")
ps["A12"] = "Riferimenti usati dalle formule"
ps["A12"].font = F_BODYB
ps["A13"] = ("Ore per giornata = Parametri!$C$5   ·   Tariffe = $C$6 / $C$7 / $C$8   ·   "
             "Lead time target = $C$10")
ps["A13"].font = F_SUB
ps.merge_cells("A13:E13")

P_ORE  = "Parametri!$C$5"
P_FRO  = "Parametri!$C$6"
P_BACK = "Parametri!$C$7"
P_TRAV = "Parametri!$C$8"
P_TGT  = "Parametri!$C$10"

# ================================================================ RUOLI
rs = wb.create_sheet("Ruoli")
rs.sheet_view.showGridLines = False
banda_titolo(rs, "A1:E1", "TEAM DI PROGETTO — RUOLI E NOMI")
rs["A2"] = ("Compilare la colonna Nome. I ruoli sono quelli richiamati nelle colonne Owner e "
            "Partecipanti della WBS.")
rs["A2"].font = F_SUB
rs.merge_cells("A2:E2")

for col, txt, w in (("A", "Ruolo", 32), ("B", "Nome", 28), ("C", "Seniority", 14),
                    ("D", "Presidio", 30), ("E", "Note", 46)):
    rs.column_dimensions[col].width = w
    rs[f"{col}4"] = txt
    rs[f"{col}4"].fill = FILL_YEL if col == "B" else FILL_BLK
    rs[f"{col}4"].font = F_HEADY if col == "B" else F_HEAD
    rs[f"{col}4"].alignment = CTR

RUOLI = [
    ("Commerciale", None, "Senior", "Fase 1 — convocazione kick off",
     "Invia la prima email al cliente per fissare il kick off."),
    ("Project Manager", None, "Senior", "Fasi 1 → 12 — presidio di progetto",
     "Invia il follow-up post kick off con documento presentato e tracciato dati. Owner del piano."),
    ("Facilitatore senior", "Alessandro Bruti", "Senior", "Tutti i workshop con il cliente",
     "Conduce i workshop di revenue model, ICP, process design, tech stack e data model."),
    ("Consulente RevOps junior 1", None, "Junior", "Workshop + back office",
     "In presenza ai workshop; analisi dati e compilazione file di delivery."),
    ("Consulente RevOps junior 2", None, "Junior", "Workshop + back office",
     "In presenza ai workshop; supporto alla delivery."),
]
for i, (ruolo, nome, sen, pres, note) in enumerate(RUOLI):
    r = 5 + i
    rs[f"A{r}"] = ruolo; rs[f"A{r}"].font = F_BODYB
    rs[f"B{r}"] = nome
    rs[f"B{r}"].font = F_IN
    rs[f"B{r}"].fill = FILL_WHT
    rs[f"B{r}"].border = INBOX
    rs[f"C{r}"] = sen; rs[f"C{r}"].font = F_BODY; rs[f"C{r}"].alignment = CTRV
    rs[f"D{r}"] = pres; rs[f"D{r}"].font = F_BODY; rs[f"D{r}"].alignment = TOP
    rs[f"E{r}"] = note; rs[f"E{r}"].font = F_BODY; rs[f"E{r}"].alignment = TOP
    rs[f"A{r}"].alignment = TOP
    rs[f"B{r}"].alignment = TOP
    for col in "ACDE":
        rs[f"{col}{r}"].border = HRULE
    rs.row_dimensions[r].height = 30
# ================================================================ DATI FASI
# (fase, tipo, modalita, blocco, sett_inizio, durata_sett, n_incontri, durata_h,
#  descrizione, prep, followup, input_cliente, output, owner, partecipanti, delivery, note)
FASI = [
 ("Kick off e raccolta dati", "Cliente", "Remoto", "1 · Avvio", 1, 1, 1, None,
  "Incontro di avvio in remoto con i team di progetto di entrambe le parti. Si condividono obiettivi e "
  "perimetro dell'assessment, si presenta il team e si calendarizzano gli incontri successivi. Nella seconda "
  "parte si illustra la struttura dei dati richiesta al cliente, si chiariscono dubbi e domande e si concorda "
  "la data di consegna degli export, fissata entro una settimana dal kick off. Il lead time della fase "
  "coincide con il tempo necessario al cliente per predisporre e trasmettere gli export.",
  "Invio dell'invito e fissaggio dell'incontro a cura del Commerciale.",
  "Invio, a cura del Project Manager, del documento presentato e del tracciato dei dati richiesti.",
  "Export dei dati trasmessi via email entro 5 giorni lavorativi dal kick off.",
  "Obiettivi e perimetro condivisi · calendario degli incontri · tracciato dati concordato.",
  "Commerciale (convocazione) → Project Manager (follow-up)",
  "Team di progetto completo, entrambe le parti",
  None,
  "Vincolo: la consegna degli export non può superare 1 settimana dal kick off."),

 ("Analisi dati e sintesi", "Interna", "Interna", "1 · Avvio", 2, 1, 1, 1,
  "Analisi degli export ricevuti per ricostruire l'evidenza sulla base clienti attuale, su due piani: lettura "
  "logico-funzionale del dato e verifica di data quality. L'esito confluisce in una restituzione interna al "
  "team, in cui si presentano le evidenze emerse, l'issue tree e le prime ipotesi da validare nel corso del "
  "progetto.",
  "Verifica di completezza e leggibilità degli export ricevuti.",
  "Formalizzazione al cliente via email di dubbi, lacune e richieste di integrazione dati.",
  "Export dati completi e leggibili.",
  "Evidenze sulla base clienti · esiti di data quality · issue tree · ipotesi da validare · lista domande al cliente.",
  "Project Manager",
  "1 consulente RevOps sull'analisi; restituzione di 1 ora all'intero team",
  "1 risorsa in back office; restituzione al team di 1 ora con tutto il gruppo",
  "L'analisi è svolta da una sola persona: la restituzione è l'unico momento collegiale."),

 ("Revenue model", "Cliente", "In presenza", "2 · Workshop in parallelo", 3, 2, 2, None,
  "Workshop in presenza per la mappatura del modello di ricavo del cliente, articolato in due sessioni "
  "distinte — una dedicata al canale B2B e una al B2C — con compilazione congiunta del framework revenue model.",
  "Predisposizione dei materiali di workshop e condivisione dell'agenda con il cliente.",
  "Compilazione del file di delivery · restituzione formalizzata al cliente · registrazione delle evidenze emerse.",
  "Disponibilità dei referenti di canale e accesso ai dati di ricavo.",
  "Revenue model mappato e validato per canale B2B e B2C.",
  "Project Manager",
  "Facilitatore senior + 2 consulenti RevOps junior",
  "3 persone in presenza (1 facilitatore senior + 2 consulenti junior)",
  "Durata delle due sessioni da definire."),

 ("ICP model e customer journey", "Cliente", "Remoto", "2 · Workshop in parallelo", 3, 2, 2, None,
  "Due incontri in remoto — uno sul canale B2B e uno sul B2C — per la definizione del profilo di cliente ideale "
  "e della relativa customer journey. La mappatura si sviluppa sul bow tie: per ciascuna fase si individua il "
  "percorso desiderato e le frizioni oggi presenti.",
  "Predisposizione della struttura bow tie e condivisione dell'agenda.",
  "Compilazione del file di delivery · restituzione formalizzata al cliente · registrazione delle evidenze emerse.",
  "Disponibilità dei referenti marketing e sales.",
  "Ideal Customer Profile definito · customer journey mappata lungo il bow tie con le frizioni per fase.",
  "Project Manager",
  "Facilitatore senior + 2 consulenti RevOps junior",
  "3 persone in remoto (1 facilitatore senior + 2 consulenti junior)",
  "Durata delle due sessioni da definire."),

 ("Process design", "Cliente", "In presenza", "2 · Workshop in parallelo", 3, 2, None, None,
  "Workshop in presenza per la mappatura dei processi attualmente in essere su marketing, sales e delivery. "
  "Si analizza cosa funziona e cosa non funziona nei flussi correnti e si individuano i gap rispetto allo stato "
  "desiderato definito nella customer journey.",
  "Raccolta della documentazione di processo esistente e condivisione dell'agenda.",
  "Compilazione del file di delivery · restituzione formalizzata al cliente · registrazione delle evidenze emerse.",
  "Documentazione di processo e disponibilità dei process owner.",
  "Mappatura del flusso di processo lungo l'intera catena del valore go-to-market · gap rispetto allo stato desiderato.",
  "Project Manager",
  "Facilitatore senior + 2 consulenti RevOps junior",
  "3 persone in presenza (1 facilitatore senior + 2 consulenti junior)",
  "Numero e durata delle sessioni da definire."),

 ("Tech stack", "Cliente", "Da definire", "2 · Workshop in parallelo", 3, 2, None, None,
  "Ricognizione degli strumenti oggi in uso che intercettano la strategia go-to-market, con verifica di come "
  "sono strutturati database e flussi di dati. L'obiettivo è fotografare lo stack as-is e disegnare lo stack "
  "to-be in ottica di razionalizzazione.",
  "Richiesta preventiva dell'inventario applicativo e delle licenze in essere.",
  "Compilazione del file di delivery · restituzione formalizzata al cliente · registrazione delle evidenze emerse.",
  "Inventario degli strumenti in uso e accesso ai referenti IT.",
  "Mappa degli strumenti as-is e to-be · indicazioni di razionalizzazione dello stack e del database.",
  "Project Manager",
  "Facilitatore senior + 2 consulenti RevOps junior",
  "3 persone (1 facilitatore senior + 2 consulenti junior)",
  "Modalità (presenza / remoto), numero e durata delle sessioni da definire."),

 ("Data model", "Cliente", "Da definire", "2 · Workshop in parallelo", 3, 2, 1, None,
  "Incontro dedicato alla proposta del set di KPI necessari a misurare il sistema, costruito utilizzando il "
  "bow tie come backbone. Si definiscono KPI di primo e di secondo livello per monitorare l'avanzamento della "
  "generazione di valore.",
  "Predisposizione dell'albero dei KPI di primo e secondo livello sul bow tie.",
  "Compilazione del file di delivery · restituzione formalizzata al cliente · registrazione delle evidenze emerse.",
  "Disponibilità dei dati necessari ad alimentare i KPI proposti.",
  "Set di KPI di primo e secondo livello ancorati al bow tie · logica di misurazione della generazione di valore.",
  "Project Manager",
  "Facilitatore senior + 2 consulenti RevOps junior",
  "3 persone (1 facilitatore senior + 2 consulenti junior)",
  "Modalità (presenza / remoto) e durata da definire."),

 ("Interviste qualitative", "Cliente", "Remoto", "3 · Approfondimento", 5, 1, None, None,
  "Ciclo di interviste individuali in remoto, condotte da un solo consulente a valle del process design. Sono "
  "interviste qualitative volte a raccogliere il punto di vista degli utenti interni sui temi emersi nei "
  "workshop e a far emergere frizioni operative non visibili nella mappatura di processo.",
  "Selezione degli intervistati con il cliente e predisposizione della guida di intervista.",
  "Sintesi trasversale delle interviste e integrazione delle frizioni nella mappatura di processo.",
  "Individuazione e disponibilità degli utenti interni da intervistare.",
  "Frizioni operative rilevate · feedback qualitativi sui temi aperti nel process design.",
  "Project Manager",
  "1 consulente RevOps; referenti operativi lato cliente",
  "1 persona in remoto",
  "Fase separata perché si svolge a valle dei workshop, con un solo consulente impegnato."),

 ("Sintesi evidenze e business case", "Interna", "Interna", "4 · Sintesi e validazione", 6, 1, None, None,
  "Fase interna di ristrutturazione di tutte le evidenze raccolte e di costruzione del piano di lavoro, con "
  "elaborazione dei business case a supporto delle iniziative individuate.",
  "Consolidamento dei file di delivery di tutte le fasi precedenti.",
  "Predisposizione dei materiali per la review con il cliente.",
  None,
  "Piano di lavoro strutturato · business case per iniziativa.",
  "Project Manager",
  "Team di progetto",
  None,
  None),

 ("Review con il cliente", "Cliente", "Remoto", "4 · Sintesi e validazione", 7, 1, 1, None,
  "Incontro di validazione con il team di lavoro lato cliente, centrato sui business case costruiti: si "
  "verifica l'allineamento su priorità e numeri e si stabilisce se procedere o iterare su specifici ambiti.",
  "Invio anticipato dei business case per lettura preventiva.",
  "Recepimento delle iterazioni richieste e aggiornamento del piano di lavoro.",
  "Presenza dei decisori lato cliente.",
  "Business case validati oppure lista delle iterazioni richieste.",
  "Project Manager",
  "Team di progetto + team di lavoro lato cliente",
  None,
  "GATE DI PROGETTO: il consolidamento del report parte solo a validazione avvenuta."),

 ("Consolidamento report finale", "Interna", "Interna", "5 · Chiusura", 8, 2, None, None,
  "Consolidamento dell'intero assessment nel report finale: stesura delle slide in versione definitiva e "
  "valorizzazione economica della roadmap di progetto.",
  "Recepimento degli esiti della review con il cliente.",
  "Condivisione anticipata del report con i referenti prima della presentazione.",
  None,
  "Report finale di assessment · roadmap di progetto valorizzata economicamente.",
  "Project Manager",
  "Team di progetto",
  None,
  None),

 ("Presentazione assessment", "Cliente", "Da definire", "5 · Chiusura", 10, 1, 1, None,
  "Presentazione al cliente del report finale di assessment, della roadmap e della relativa valorizzazione "
  "economica, come passaggio di consegna verso la fase successiva.",
  "Allineamento sull'agenda e sui partecipanti lato cliente.",
  "Consegna formale del report e apertura del confronto sulla fase successiva.",
  "Presenza dei decisori lato cliente.",
  "Assessment consegnato · roadmap valorizzata condivisa · next step definiti.",
  "Project Manager",
  "Team di progetto + decisori lato cliente",
  None,
  "Modalità (presenza / remoto) da definire."),
]


# ================================================================ WBS
ws = wb.create_sheet("WBS", 0)
ws.sheet_view.showGridLines = False

COLS = [
    ("A", "#", 5,  "num"),
    ("B", "Cod. WBS", 10, "txt"),
    ("C", "Fase", 30, "txt"),
    ("D", "Tipo", 11, "txt"),
    ("E", "Modalità", 13, "txt"),
    ("F", "Blocco", 24, "txt"),
    ("G", "Sett.\ninizio", 8, "in"),
    ("H", "Durata\n(sett.)", 8, "in"),
    ("I", "Sett.\nfine", 8, "calc"),
    ("J", "Descrizione", 68, "txt"),
    ("K", "Attività preparatorie", 40, "txt"),
    ("L", "Attività di follow-up", 40, "txt"),
    ("M", "Input dal cliente", 34, "txt"),
    ("N", "Output / Deliverable", 44, "txt"),
    ("O", "Owner", 26, "txt"),
    ("P", "Partecipanti", 30, "txt"),
    ("Q", "N.\nincontri", 8, "in"),
    ("R", "Durata\nincontro (h)", 10, "in"),
    ("S", "gg Front\noffice", 10, "in"),
    ("T", "gg Back\noffice", 10, "in"),
    ("U", "gg\nViaggio", 9, "in"),
    ("V", "Delivery (composizione team)", 38, "txt"),
    ("W", "Tot.\ngiornate", 10, "calc"),
    ("X", "Ore\ntotali", 10, "calc"),
    ("Y", "Costo totale", 15, "calc"),
    ("Z", "Note", 46, "txt"),
]
LAST = "Z"
HR, R0 = 5, 6
R1 = R0 + len(FASI) - 1
RT = R1 + 1

banda_titolo(ws, f"A1:{LAST}1", "WORK BREAKDOWN STRUCTURE — PROGETTO DI ASSESSMENT GO-TO-MARKET")

ws.merge_cells(f"A2:{LAST}2")
ws["A2"] = ("LEGENDA — Le colonne con intestazione GIALLA sono da compilare a mano: le loro celle "
            "portano un riquadro nero. Le celle su fondo GRIGIO sono calcolate da formula e non vanno "
            "modificate. Tariffe e ore per giornata si impostano nel foglio «Parametri», i nomi del "
            "team nel foglio «Ruoli».")
ws["A2"].font = F_SUB; ws["A2"].alignment = TOP
ws.row_dimensions[2].height = 24

ws.merge_cells(f"A3:{LAST}3")
ws["A3"] = ("Le colonne «gg Front office / Back office / Viaggio» esprimono GIORNATE-UOMO aggregate, "
            "non giorni di calendario: 2 consulenti per mezza giornata dal cliente = 1 giornata front. "
            "La composizione della squadra si descrive nella colonna «Delivery». Le ore di viaggio sono "
            "incluse nel monte ore.")
ws["A3"].font = F_ACC
ws["A3"].alignment = TOP
for i in range(ord("A"), ord(LAST) + 1):
    ws[f"{chr(i)}3"].fill = FILL_YEL
ws.row_dimensions[3].height = 26

for letter, header, width, kind in COLS:
    ws.column_dimensions[letter].width = width
    c = ws[f"{letter}{HR}"]
    c.value = header
    c.fill = FILL_YEL if kind == "in" else FILL_BLK
    c.font = F_HEADY if kind == "in" else F_HEAD
    c.alignment = CTR
ws.row_dimensions[HR].height = 34

for i, f in enumerate(FASI):
    (fase, tipo, mod, blocco, w_start, w_dur, n_inc, dur_h,
     descr, prep, follow, inp, outp, owner, part, deliv, note) = f
    r = R0 + i
    vals = {
        "A": i + 1, "B": f"1.{i+1}",
        "C": fase, "D": tipo, "E": mod, "F": blocco,
        "G": w_start, "H": w_dur,
        "I": f'=IF(AND(G{r}<>"",H{r}<>""),G{r}+H{r}-1,"")',
        "J": descr, "K": prep, "L": follow, "M": inp, "N": outp, "O": owner, "P": part,
        "Q": n_inc, "R": dur_h, "S": None, "T": None, "U": None, "V": deliv,
        "W": f'=IF(COUNT(S{r}:U{r})=0,"",SUM(S{r}:U{r}))',
        "X": f'=IF(W{r}="","",W{r}*{P_ORE})',
        "Y": (f'=IF(COUNT(S{r}:U{r})=0,"",S{r}*{P_ORE}*{P_FRO}'
              f'+T{r}*{P_ORE}*{P_BACK}+U{r}*{P_ORE}*{P_TRAV})'),
        "Z": note,
    }
    for letter, _, _, kind in COLS:
        c = ws[f"{letter}{r}"]
        c.value = vals[letter]
        if kind == "in":
            c.font = F_IN; c.fill = FILL_WHT; c.border = INBOX
            c.alignment = CTRV; c.number_format = NUM
        elif kind == "calc":
            c.font = F_CALC; c.fill = FILL_S2; c.border = HRULE
            c.alignment = CTRV
            c.number_format = EUR if letter == "Y" else NUM
        elif kind == "num":
            c.font = F_BODYB; c.alignment = CTRV; c.border = HRULE
        else:
            c.font = F_BODYB if letter == "C" else F_BODY
            c.alignment = TOP; c.border = HRULE
    ws.row_dimensions[r].height = 78

# --- riga totale: fondo nero, testo bianco, filetto giallo sopra
ws.merge_cells(f"A{RT}:R{RT}")
ws[f"A{RT}"] = "TOTALE PROGETTO"
ws[f"A{RT}"].font = F_TOTAL
ws[f"A{RT}"].alignment = Alignment(horizontal="right", vertical="center")
for letter, _, _, _ in COLS:
    c = ws[f"{letter}{RT}"]
    c.fill = FILL_BLK; c.font = F_TOTAL; c.alignment = CTR; c.border = OVERY
for letter in ("S", "T", "U", "W", "X", "Y"):
    ws[f"{letter}{RT}"] = f"=SUM({letter}{R0}:{letter}{R1})"
    ws[f"{letter}{RT}"].number_format = EUR if letter == "Y" else NUM
ws.row_dimensions[RT].height = 22

# --- riga di esempio, fuori dal totale
RE = RT + 2
RX = RE + 1
ws.merge_cells(f"A{RE}:{LAST}{RE}")
ws[f"A{RE}"] = ("ESEMPIO DI COMPILAZIONE — riga dimostrativa, NON conteggiata nei totali. "
                "Mostra il formato atteso per le colonne numeriche.")
ws[f"A{RE}"].font = Font(name=FONT, size=9, bold=True, italic=True, color=INK2)
ws[f"A{RE}"].alignment = TOP
for i in range(ord("A"), ord(LAST) + 1):
    ws[f"{chr(i)}{RE}"].fill = FILL_S2

ex = {
    "A": "es.", "B": "1.x", "C": "Nome della fase", "D": "Cliente", "E": "In presenza",
    "F": "2 · Workshop in parallelo", "G": 3, "H": 2, "I": f'=G{RX}+H{RX}-1',
    "J": "Descrizione sintetica di cosa avviene nella fase.",
    "K": "Cosa si fa prima.", "L": "Cosa si fa dopo.", "M": "Cosa serve dal cliente.",
    "N": "Cosa esce dalla fase.", "O": "Project Manager",
    "P": "Facilitatore senior + 2 consulenti junior",
    "Q": 2, "R": 4, "S": 1.5, "T": 2, "U": 0.5,
    "V": "3 persone in presenza, 1 sola lavora in back office",
    "W": f'=SUM(S{RX}:U{RX})', "X": f'=W{RX}*{P_ORE}',
    "Y": f'=S{RX}*{P_ORE}*{P_FRO}+T{RX}*{P_ORE}*{P_BACK}+U{RX}*{P_ORE}*{P_TRAV}',
    "Z": "1,5 gg front = 3 consulenti per mezza giornata + 3 consulenti per mezza giornata.",
}
for letter, _, _, kind in COLS:
    c = ws[f"{letter}{RX}"]
    c.value = ex[letter]
    c.font = F_EX; c.fill = FILL_S2; c.border = HRULE
    c.alignment = CTRV if kind in ("in", "calc", "num") else TOP
    if kind == "calc":
        c.number_format = EUR if letter == "Y" else NUM
    elif kind == "in":
        c.number_format = NUM
ws.row_dimensions[RX].height = 44

dv_tipo = DataValidation(type="list", formula1='"Cliente,Interna"', allow_blank=True)
dv_mod  = DataValidation(type="list", formula1='"In presenza,Remoto,Interna,Da definire"', allow_blank=True)
ws.add_data_validation(dv_tipo); ws.add_data_validation(dv_mod)
dv_tipo.add(f"D{R0}:D{R1}")
dv_mod.add(f"E{R0}:E{R1}")

ws.freeze_panes = f"D{R0}"
ws.auto_filter.ref = f"A{HR}:{LAST}{R1}"

# ================================================================ TIMELINE
ts = wb.create_sheet("Timeline", 1)
ts.sheet_view.showGridLines = False
NW, W_FIRST = 12, 7
THR, RMK, T0 = 4, 5, 6
T1 = T0 + len(FASI) - 1
RA = T1 + 1
W_LAST = get_column_letter(W_FIRST + NW - 1)
GC = get_column_letter(W_FIRST)

banda_titolo(ts, f"A1:{W_LAST}1", "TIMELINE DI PROGETTO — SOVRAPPOSIZIONE DELLE FASI PER SETTIMANA")

ts.merge_cells(f"A2:{W_LAST}2")
ts["A2"] = ("Le barre si disegnano da sole dalle settimane impostate nella WBS: qui non si scrive nulla. "
            "GIALLO = le fasi che si svolgono nelle STESSE settimane, in parallelo. NERO = altra fase con "
            "il cliente. GRIGIO = fase interna. La riga in fondo conta quante fasi sono attive in "
            "ciascuna settimana.")
ts["A2"].font = F_SUB; ts["A2"].alignment = TOP
ts.row_dimensions[2].height = 24

THEAD = [("A", "#", 5), ("B", "Fase", 32), ("C", "Tipo", 11),
         ("D", "Blocco", 24), ("E", "Da\nsett.", 8), ("F", "A\nsett.", 8)]
for letter, txt, w in THEAD:
    ts.column_dimensions[letter].width = w
    c = ts[f"{letter}{THR}"]
    c.value = txt; c.font = F_HEAD; c.fill = FILL_BLK; c.alignment = CTR
for k in range(NW):
    letter = get_column_letter(W_FIRST + k)
    ts.column_dimensions[letter].width = 5.5
    c = ts[f"{letter}{THR}"]
    c.value = k + 1
    c.number_format = '"S"0'
    c.font = F_HEAD; c.fill = FILL_BLK; c.alignment = CTR
ts.row_dimensions[THR].height = 24

# riga marcatore: chevron giallo sulle settimane con più di una fase attiva
ts.merge_cells(f"A{RMK}:F{RMK}")
ts[f"A{RMK}"] = "Settimane con fasi in parallelo  ›"
ts[f"A{RMK}"].font = F_BODYB
ts[f"A{RMK}"].alignment = Alignment(horizontal="right", vertical="center")
for k in range(NW):
    letter = get_column_letter(W_FIRST + k)
    c = ts[f"{letter}{RMK}"]
    c.value = f'=IF({letter}{RA}>1,"›","")'
    c.font = F_ACC; c.alignment = CTR
ts.row_dimensions[RMK].height = 18
ts.conditional_formatting.add(f"{GC}{RMK}:{W_LAST}{RMK}", FormulaRule(
    formula=[f'{GC}{RMK}<>""'], fill=FILL_YEL,
    font=Font(name=FONT, size=9, bold=True, color=BLK), stopIfTrue=True))

for i in range(len(FASI)):
    r, wr = T0 + i, R0 + i
    ts[f"A{r}"] = f"=WBS!A{wr}"
    ts[f"B{r}"] = f"=WBS!C{wr}"
    ts[f"C{r}"] = f"=WBS!D{wr}"
    ts[f"D{r}"] = f"=WBS!F{wr}"
    ts[f"E{r}"] = f"=WBS!G{wr}"
    ts[f"F{r}"] = f"=WBS!I{wr}"
    for letter in ("A", "B", "C", "D", "E", "F"):
        c = ts[f"{letter}{r}"]
        c.font = F_CALCB if letter == "B" else F_CALC
        c.border = HRULE
        c.alignment = TOP if letter in ("B", "D") else CTRV
    for k in range(NW):
        letter = get_column_letter(W_FIRST + k)
        c = ts[f"{letter}{r}"]
        c.value = (f'=IF(AND($E{r}<>"",$F{r}<>"",$E{r}<={letter}${THR},'
                   f'$F{r}>={letter}${THR}),1,"")')
        c.font = F_BODY
        c.number_format = ";;;"      # il valore resta invisibile: si vede solo il pieno
        c.border = HRULE
        c.alignment = CTR
    ts.row_dimensions[r].height = 22

# scala monocroma + unico accento: il giallo va alla serie che porta il messaggio
GRID = f"{GC}{T0}:{W_LAST}{T1}"
ts.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'AND(ISNUMBER(SEARCH("parallelo",$D{T0})),{GC}{T0}=1)'],
    fill=FILL_YEL, stopIfTrue=True))
ts.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'AND($C{T0}="Interna",{GC}{T0}=1)'],
    fill=FILL_GMID, stopIfTrue=True))
ts.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'{GC}{T0}=1'], fill=FILL_BLK, stopIfTrue=True))

# le fasi del blocco parallelo si distinguono anche per peso del carattere
ts.conditional_formatting.add(f"B{T0}:F{T1}", FormulaRule(
    formula=[f'ISNUMBER(SEARCH("parallelo",$D{T0}))'],
    font=Font(name=FONT, size=9, bold=True, color=BLK)))

# --- riga di conteggio
ts.merge_cells(f"A{RA}:F{RA}")
ts[f"A{RA}"] = "Fasi attive nella settimana"
ts[f"A{RA}"].font = F_TOTAL
ts[f"A{RA}"].alignment = Alignment(horizontal="right", vertical="center")
for letter in ("A", "B", "C", "D", "E", "F"):
    ts[f"{letter}{RA}"].fill = FILL_BLK
    ts[f"{letter}{RA}"].border = OVERY
for k in range(NW):
    letter = get_column_letter(W_FIRST + k)
    c = ts[f"{letter}{RA}"]
    c.value = f"=COUNT({letter}{T0}:{letter}{T1})"
    c.font = F_TOTAL; c.fill = FILL_BLK; c.alignment = CTR
    c.border = OVERY; c.number_format = INT
ts.row_dimensions[RA].height = 22
ts.conditional_formatting.add(f"{GC}{RA}:{W_LAST}{RA}", FormulaRule(
    formula=[f'{GC}{RA}>1'], fill=FILL_YEL,
    font=Font(name=FONT, size=9, bold=True, color=BLK), stopIfTrue=True))

RB = RA + 2
for r, label, formula in (
    (RB,     "Settimane con più di una fase attiva in parallelo", f'=COUNTIF({GC}{RA}:{W_LAST}{RA},">1")'),
    (RB + 1, "Settimana di massima concentrazione (n. fasi in parallelo)", f'=MAX({GC}{RA}:{W_LAST}{RA})'),
):
    ts[f"A{r}"] = label
    ts[f"A{r}"].font = F_BODYB
    ts.merge_cells(f"A{r}:E{r}")
    c = ts[f"F{r}"]
    c.value = formula
    c.font = F_CALCB; c.fill = FILL_S2; c.border = CALCBOX
    c.alignment = CTR; c.number_format = INT

ts.freeze_panes = f"{GC}{T0}"

# ================================================================ RIEPILOGO
sm = wb.create_sheet("Riepilogo", 2)
sm.sheet_view.showGridLines = False
for col, w in (("A", 58), ("B", 2), ("C", 18), ("D", 13), ("E", 62)):
    sm.column_dimensions[col].width = w

banda_titolo(sm, "A1:E1", "RIEPILOGO DI PROGETTO")
sm.merge_cells("A2:E2")
sm["A2"] = ("Tutti i valori sono calcolati dalla WBS e dai Parametri. Nessuna cella di questo foglio "
            "va compilata a mano.")
sm["A2"].font = F_SUB
sm.row_dimensions[2].height = 18


def sect(r, title):
    """Fascia di sezione: superficie d'accento — fondo giallo, testo nero."""
    sm.merge_cells(f"A{r}:E{r}")
    sm[f"A{r}"] = f"›  {title}"
    sm[f"A{r}"].font = Font(name=FONT, size=10, bold=True, color=BLK)
    sm[f"A{r}"].alignment = Alignment(horizontal="left", vertical="center")
    for col in "ABCDE":
        sm[f"{col}{r}"].fill = FILL_YEL
    sm.row_dimensions[r].height = 20


def line(r, label, formula, fmt=NUM, unit="", note="", bold=False):
    sm[f"A{r}"] = label
    sm[f"A{r}"].font = F_BODYB if bold else F_BODY
    sm[f"A{r}"].alignment = TOP
    sm[f"A{r}"].border = HRULE
    c = sm[f"C{r}"]
    c.value = formula
    c.font = F_CALCB if bold else F_CALC
    c.fill = FILL_S2; c.border = HRULE; c.alignment = CTRV; c.number_format = fmt
    sm[f"D{r}"] = unit; sm[f"D{r}"].font = F_BODY
    sm[f"D{r}"].alignment = CTRV; sm[f"D{r}"].border = HRULE
    sm[f"E{r}"] = note; sm[f"E{r}"].font = F_SUB
    sm[f"E{r}"].alignment = TOP; sm[f"E{r}"].border = HRULE
    return c


D_RNG = f"WBS!$D${R0}:$D${R1}"
W_RNG = f"WBS!$W${R0}:$W${R1}"
Y_RNG = f"WBS!$Y${R0}:$Y${R1}"
Q_RNG = f"WBS!$Q${R0}:$Q${R1}"

sect(4, "IMPEGNO E COSTO COMPLESSIVO")
line(5, "Giornate front office (presso il cliente)", f"=WBS!S{RT}", NUM, "gg-uomo")
line(6, "Giornate back office", f"=WBS!T{RT}", NUM, "gg-uomo")
line(7, "Giornate di viaggio", f"=WBS!U{RT}", NUM, "gg-uomo",
     "Incluse nel monte ore, come da impostazione concordata.")
line(8, "TOTALE giornate-uomo di progetto", f"=WBS!W{RT}", NUM, "gg-uomo", "", bold=True)
line(9, "TOTALE ore di progetto", f"=WBS!X{RT}", NUM, "h",
     f"Totale giornate × ore per giornata ({P_ORE}).", bold=True)
line(10, "COSTO TOTALE DI PROGETTO", f"=WBS!Y{RT}", EUR, "€", "", bold=True)

sect(12, "RIPARTIZIONE DEL COSTO PER TIPOLOGIA DI ORA")
line(13, "Costo ore front office", f"=WBS!S{RT}*{P_ORE}*{P_FRO}", EUR, "€")
line(14, "Costo ore back office", f"=WBS!T{RT}*{P_ORE}*{P_BACK}", EUR, "€")
line(15, "Costo ore di viaggio", f"=WBS!U{RT}*{P_ORE}*{P_TRAV}", EUR, "€")
line(16, "Quadratura (deve coincidere con il costo totale)", "=SUM(C13:C15)", EUR, "€",
     "Controllo interno: se differisce dal costo totale c'è un errore di formula.", bold=True)

sect(18, "RIPARTIZIONE PER TIPO DI FASE")
line(19, "Giornate — fasi con il cliente", f'=SUMIF({D_RNG},"Cliente",{W_RNG})', NUM, "gg-uomo")
line(20, "Giornate — fasi interne", f'=SUMIF({D_RNG},"Interna",{W_RNG})', NUM, "gg-uomo")
line(21, "Costo — fasi con il cliente", f'=SUMIF({D_RNG},"Cliente",{Y_RNG})', EUR, "€")
line(22, "Costo — fasi interne", f'=SUMIF({D_RNG},"Interna",{Y_RNG})', EUR, "€")
line(23, "Incontri previsti con il cliente", f'=SUMIF({D_RNG},"Cliente",{Q_RNG})', INT, "n.")
line(24, "Fasi con il cliente / fasi interne",
     f'=COUNTIF({D_RNG},"Cliente")&" / "&COUNTIF({D_RNG},"Interna")', "General", "")

sect(26, "LEAD TIME E EFFETTO DELLA PARALLELIZZAZIONE")
line(27, "Somma delle durate delle fasi, se svolte in sequenza",
     f"=SUM(WBS!H{R0}:H{R1})", INT, "settimane",
     "Somma aritmetica della colonna «Durata (sett.)» della WBS.")
line(28, "Lead time di calendario effettivo",
     f"=MAX(WBS!I{R0}:I{R1})-MIN(WBS!G{R0}:G{R1})+1", INT, "settimane",
     "Prima settimana di inizio → ultima settimana di fine, tenendo conto delle fasi sovrapposte.",
     bold=True)
line(29, "Settimane recuperate dalla parallelizzazione", "=C27-C28", INT, "settimane",
     "Effetto dei workshop svolti nelle stesse due settimane.")
line(30, "Lead time target dichiarato", f"={P_TGT}", INT, "settimane",
     "Impostato nel foglio «Parametri».")
line(31, "SCOSTAMENTO effettivo vs target", "=C28-C30", INT, "settimane",
     "Zero = il piano quadra. Se lo scostamento non è zero la cella si inverte in nero.", bold=True)

# Nessun colore semantico: lo stato è codificato con la scala monocroma
# (fondo nero, testo bianco) e con un'etichetta in parole.
sm.conditional_formatting.add("C31", FormulaRule(
    formula=["C31<>0"], fill=FILL_BLK,
    font=Font(name=FONT, size=9, bold=True, color=WHT), stopIfTrue=True))

sm["A32"] = "Esito del controllo"
sm["A32"].font = F_BODYB; sm["A32"].border = HRULE
sm.merge_cells("C32:E32")
sm["C32"] = ('=IF(C28=C30,"Il piano quadra con il lead time target.",'
             '"NON QUADRA — rivedere le settimane di inizio e le durate nella WBS.")')
sm["C32"].font = F_BODY
sm["C32"].alignment = Alignment(horizontal="left", vertical="center")
sm.conditional_formatting.add("C32:E32", FormulaRule(
    formula=["$C$28<>$C$30"], fill=FILL_BLK,
    font=Font(name=FONT, size=9, bold=True, color=WHT), stopIfTrue=True))
sm.row_dimensions[32].height = 18

sect(34, "IPOTESI E PUNTI APERTI")
NOTES = [
 ("Pianificazione delle settimane",
  "Il piano proposto è: S1 kick off · S2 analisi dati · S3–S4 i cinque workshop con il cliente in "
  "parallelo (revenue model, ICP e customer journey, process design, tech stack, data model) · S5 "
  "interviste qualitative a valle del process design · S6 sintesi e business case · S7 review con il "
  "cliente · S8–S9 consolidamento report · S10 presentazione. Totale 10 settimane, coerente con il "
  "target dichiarato."),
 ("Giornate front / back / viaggio",
  "Non ancora fornite: le celle sono vuote, sotto intestazione gialla. Finché non sono compilate, ore "
  "e costi restano a zero."),
 ("Tariffe orarie",
  "Non ancora fornite: da inserire nel foglio «Parametri». Sono tre tariffe distinte (front, back, "
  "viaggio)."),
 ("Durata dei workshop",
  "Da definire per revenue model, ICP e customer journey, process design, tech stack, data model e "
  "presentazione finale (colonna «Durata incontro (h)» della WBS)."),
 ("Modalità da confermare",
  "Tech stack, data model e presentazione finale sono impostati su «Da definire»: da scegliere fra "
  "presenza e remoto. La scelta incide sulle giornate di viaggio."),
 ("Somma sequenziale vs lead time",
  "La somma delle durate di fase vale 18 settimane perché le cinque fasi in parallelo occupano "
  "ciascuna la stessa finestra di 2 settimane. Il lead time reale resta 10 settimane: è questo il "
  "numero da leggere."),
 ("Identità visiva",
  "Palette Impresoft di tre colori (giallo #FDC300, nero, bianco) più grigi puri derivati dal nero. "
  "Font Manrope, fallback Arial. Nessun colore semantico: verde e rosso sono stati sostituiti dalla "
  "scala monocroma e da etichette in parole."),
]
r = 35
for titolo, testo in NOTES:
    sm[f"A{r}"] = titolo
    sm[f"A{r}"].font = F_BODYB; sm[f"A{r}"].alignment = TOP; sm[f"A{r}"].border = HRULE
    sm.merge_cells(f"C{r}:E{r}")
    sm[f"C{r}"] = testo
    sm[f"C{r}"].font = F_BODY; sm[f"C{r}"].alignment = TOP
    sm.row_dimensions[r].height = 44
    r += 1

# ================================================================ RIFINITURE
for r in list(range(R0, R1 + 1)) + [RX]:
    for letter in ("G", "H", "I", "Q"):
        ws[f"{letter}{r}"].number_format = INT
ws[f"Q{RT}"].number_format = INT
for r in range(T0, T1 + 1):
    for letter in ("A", "E", "F"):
        ts[f"{letter}{r}"].number_format = INT

for sh, land in ((ws, True), (ts, True), (sm, False), (ps, False), (rs, False)):
    sh.page_setup.orientation = "landscape" if land else "portrait"
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr.fitToPage = True
    sh.print_title_rows = f"{HR}:{HR}" if sh is ws else None

wb.save(OUT)
print("scritto:", OUT)
