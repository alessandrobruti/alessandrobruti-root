# -*- coding: utf-8 -*-
"""
WBS del progetto RevOps Assessment B2B e B2C per Laica S.p.A.
Allineata all'Offerta CX260803_Off_Laica_RevOps assessment B2B e B2C v1.1 (28 ago 2026).

Tutti i calcoli sono formule Excel vive: il file si ricalcola all'apertura.
Le giornate front office e viaggio sono CALCOLATE da n. incontri × durata × n. persone;
le giornate back office sono l'unico input di effort, e sono stime Impresoft — il
contratto non quantifica il back office in nessun punto.

Identità visiva: Impresoft Brand Manual & Corporate Guidelines (agg. 27.04.2023),
ricetta Excel di references/office-docs.md. Palette di tre colori più grigi puri
derivati dal nero. Nessun colore semantico.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment

OUT = "/home/user/alessandrobruti-root/wbs-assessment/WBS_Assessment_GTM.xlsx"

# ---------------------------------------------------------------- brand
# Font istituzionale. Se Manrope non è installato sulle macchine che aprono
# il file, sostituire con "Arial" (fallback indicato dal manuale).
FONT = "Manrope"

YEL, BLK, WHT = "FDC300", "000000", "FFFFFF"
S2, INK2, INK3 = "F4F4F4", "3A3A3A", "6E6E6E"
RULE, RULS, G_MID = "DCDCDC", "9B9B9B", "8C8C8C"

F_TITLE = Font(name=FONT, size=13, bold=True, color=WHT)
F_SUB   = Font(name=FONT, size=8.5, color=INK3)
F_HEAD  = Font(name=FONT, size=9, bold=True, color=WHT)
F_HEADY = Font(name=FONT, size=9, bold=True, color=BLK)
F_BODY  = Font(name=FONT, size=9, color=BLK)
F_BODYB = Font(name=FONT, size=9, bold=True, color=BLK)
F_IN    = Font(name=FONT, size=9, color=BLK)
F_CALC  = Font(name=FONT, size=9, color=INK2)
F_CALCB = Font(name=FONT, size=9, bold=True, color=INK2)
F_TOTAL = Font(name=FONT, size=9, bold=True, color=WHT)
F_EX    = Font(name=FONT, size=9, italic=True, color=INK3)
F_ACC   = Font(name=FONT, size=9, bold=True, color=BLK)


# I fill vanno dichiarati in ARGB a 8 cifre con entrambi i colori: openpyxl
# omette l'attributo quando il valore coincide con il proprio default, e
# "000000" normalizzato È quel default — un fill nero dichiarato in modo
# abbreviato sparisce dai dxf della formattazione condizionale.
def fill(hex6):
    argb = "FF" + hex6
    return PatternFill("solid", start_color=argb, end_color=argb)


FILL_BLK, FILL_YEL = fill(BLK), fill(YEL)
FILL_S2, FILL_WHT, FILL_GMID = fill(S2), fill(WHT), fill(G_MID)

r_thin  = Side(style="thin", color=RULE)
r_black = Side(style="thin", color=BLK)
r_grey  = Side(style="thin", color=RULS)
r_yel   = Side(style="thick", color=YEL)

HRULE   = Border(bottom=r_thin)
INBOX   = Border(left=r_black, right=r_black, top=r_black, bottom=r_black)
CALCBOX = Border(left=r_grey, right=r_grey, top=r_grey, bottom=r_grey)
UNDERY  = Border(bottom=r_yel)
OVERY   = Border(top=r_yel)

TOP  = Alignment(horizontal="left",   vertical="top", wrap_text=True)
CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)
CTRV = Alignment(horizontal="center", vertical="top")
LFTC = Alignment(horizontal="left",   vertical="center")
RGTC = Alignment(horizontal="right",  vertical="center")

EUR = '#,##0.00\\ "€";-#,##0.00\\ "€";"-"'
NUM = '#,##0.00;-#,##0.00;"-"'
INT = '#,##0;-#,##0;"-"'

wb = Workbook()


def banda(sh, row, c_from, c_to, testo, fill_bg=None, font=None, height=26, brd=None):
    """Fascia a piena larghezza. c_from/c_to sono lettere di colonna."""
    i0, i1 = column_index_from_string(c_from), column_index_from_string(c_to)
    sh.merge_cells(start_row=row, start_column=i0, end_row=row, end_column=i1)
    c = sh.cell(row=row, column=i0)
    c.value = testo
    c.font = font or F_TITLE
    c.alignment = LFTC
    for i in range(i0, i1 + 1):
        cc = sh.cell(row=row, column=i)
        cc.fill = fill_bg or FILL_BLK
        if brd is not None:
            cc.border = brd
    sh.row_dimensions[row].height = height


def banda_titolo(sh, row, c_from, c_to, testo):
    banda(sh, row, c_from, c_to, f"›  {testo}", FILL_BLK, F_TITLE, 26, UNDERY)


# ================================================================ PARAMETRI
ps = wb.create_sheet("Parametri")
ps.sheet_view.showGridLines = False
banda_titolo(ps, 1, "A", "E", "PARAMETRI DI CALCOLO")
ps["A2"] = ("Le celle con riquadro nero, sotto l'intestazione gialla, sono da compilare. Tutto il file si "
            "ricalcola da qui: giornate front office e viaggio, ore, costi e costo delle trasferte.")
ps["A2"].font = F_SUB
ps.merge_cells("A2:E2")
ps.row_dimensions[2].height = 24

for col, txt, w in (("A", "Parametro", 46), ("B", "", 2), ("C", "Valore", 14),
                    ("D", "Unità", 13), ("E", "Note", 62)):
    ps.column_dimensions[col].width = w
    ps[f"{col}4"].fill = FILL_YEL if col == "C" else FILL_BLK
    if txt:
        ps[f"{col}4"] = txt
        ps[f"{col}4"].font = F_HEADY if col == "C" else F_HEAD
        ps[f"{col}4"].alignment = CTR

SEZ = "sez"
PARAMS = [
    (SEZ, 5, "IMPEGNO E TARIFFE"),
    (6,  "Ore per giornata", 8, INT, "h",
     "Base di conversione giornate → ore."),
    (7,  "Tariffa oraria — Front office", None, EUR, "€/h",
     "Ore erogate al cliente, incontri in remoto inclusi."),
    (8,  "Tariffa oraria — Back office", None, EUR, "€/h",
     "Lavoro interno: analisi, delivery, reportistica."),
    (9,  "Tariffa oraria — Viaggio", None, EUR, "€/h",
     "Tempo di trasferimento. Incluso nel monte ore di progetto."),
    (10, "Ore di viaggio per trasferta (a/r, per persona)", 1, NUM, "h",
     "Vicenza ↔ Barbarano Mossano, ~25 km per tratta. STIMA, non contrattuale."),
    (SEZ, 12, "PROGETTO"),
    (13, "Lead time target di progetto", 10, INT, "settimane",
     "Offerta § 8. Il § 6 dichiara invece 8-10 settimane: incoerenza interna al contratto."),
    (14, "Investimento contrattuale (una tantum)", 27000, EUR, "€",
     "Offerta § 10. Comprende già i costi di trasferta dei 6 incontri in presenza (§ 10.1)."),
    (SEZ, 16, "COSTO DELLE TRASFERTE (già incluso nell'investimento)"),
    (17, "Km andata e ritorno per trasferta", 50, INT, "km",
     "STIMA sulle sedi indicate al § 10.1 (Barbarano Mossano ↔ Vicenza)."),
    (18, "Ammortamento ACI", 0.65, EUR, "€/km",
     "Tariffa indicata al § 10.1 dell'Offerta per le trasferte fuori perimetro."),
    (19, "Pedaggi — casello", 0, EUR, "€/auto/trasferta",
     "Tratta ordinaria: default 0. Da correggere se il percorso prevede autostrada."),
    (20, "Vitto", None, EUR, "€/persona/trasferta",
     "Da compilare."),
    (21, "Persone per auto", 3, INT, "n.",
     "Quante persone viaggiano insieme: determina il numero di viaggi auto."),
]
for row in PARAMS:
    if row[0] == SEZ:
        _, r, title = row
        banda(ps, r, "A", "E", f"›  {title}", FILL_YEL, Font(name=FONT, size=10, bold=True, color=BLK), 20)
        continue
    r, label, val, fmt, unit, note = row
    ps[f"A{r}"] = label; ps[f"A{r}"].font = F_BODYB; ps[f"A{r}"].alignment = TOP
    c = ps[f"C{r}"]
    c.value = val
    c.font = F_IN; c.fill = FILL_WHT; c.border = INBOX
    c.alignment = CTRV; c.number_format = fmt
    ps[f"D{r}"] = unit; ps[f"D{r}"].font = F_BODY; ps[f"D{r}"].alignment = CTRV
    ps[f"E{r}"] = note; ps[f"E{r}"].font = F_BODY; ps[f"E{r}"].alignment = TOP
    for col in "ABDE":
        ps[f"{col}{r}"].border = HRULE
    ps.row_dimensions[r].height = 26

ps["C7"].comment = Comment("Le tre tariffe orarie non sono state fornite. Finché restano vuote, "
                           "tutte le colonne di costo valgono 0.", "WBS")

P_ORE   = "Parametri!$C$6"
P_FRO   = "Parametri!$C$7"
P_BACK  = "Parametri!$C$8"
P_TRAV  = "Parametri!$C$9"
P_HVIAG = "Parametri!$C$10"
P_TGT   = "Parametri!$C$13"
P_INV   = "Parametri!$C$14"
P_KM    = "Parametri!$C$17"
P_ACI   = "Parametri!$C$18"
P_PED   = "Parametri!$C$19"
P_VITTO = "Parametri!$C$20"
P_PAUTO = "Parametri!$C$21"

# ================================================================ RUOLI
rs = wb.create_sheet("Ruoli")
rs.sheet_view.showGridLines = False
banda_titolo(rs, 1, "A", "E", "TEAM DI PROGETTO — RUOLI E NOMI")
rs["A2"] = ("Compilare la colonna Nome. Sulla WBS le persone sono indicate come «facilitatore senior» e "
            "«RevOps Consultant junior» senza specializzazione funzionale, per scelta di pianificazione.")
rs["A2"].font = F_SUB
rs.merge_cells("A2:E2")
rs.row_dimensions[2].height = 24

for col, txt, w in (("A", "Ruolo", 32), ("B", "Nome", 28), ("C", "Seniority", 14),
                    ("D", "Presidio", 34), ("E", "Note", 52)):
    rs.column_dimensions[col].width = w
    rs[f"{col}4"] = txt
    rs[f"{col}4"].fill = FILL_YEL if col == "B" else FILL_BLK
    rs[f"{col}4"].font = F_HEADY if col == "B" else F_HEAD
    rs[f"{col}4"].alignment = CTR

RUOLI = [
    ("Commerciale", None, "Senior", "Fase 1 — convocazione del kick-off",
     "Nell'Offerta il responsabile è il Sales Manager."),
    ("Project Manager", None, "Senior", "Fasi 1 → 11 — presidio di progetto",
     "Referente unico verso il Cliente (Offerta § 7). Owner del piano di lavoro."),
    ("Facilitatore senior", "Alessandro Bruti", "Senior", "Tutte le sessioni con il cliente",
     "Conduce i workshop. Nell'Offerta corrisponde al RevOps Consultant senior."),
    ("RevOps Consultant junior 1", None, "Junior", "Sessioni con il cliente + back office",
     "Nell'Offerta le figure specialistiche sono CX Strategist, Process Analyst e Integration Expert."),
    ("RevOps Consultant junior 2", None, "Junior", "Sessioni con il cliente + back office",
     "Conduce le 5 interviste one-to-one della fase 5."),
]
for i, (ruolo, nome, sen, pres, note) in enumerate(RUOLI):
    r = 5 + i
    rs[f"A{r}"] = ruolo; rs[f"A{r}"].font = F_BODYB; rs[f"A{r}"].alignment = TOP
    rs[f"B{r}"] = nome
    rs[f"B{r}"].font = F_IN; rs[f"B{r}"].fill = FILL_WHT
    rs[f"B{r}"].border = INBOX; rs[f"B{r}"].alignment = TOP
    rs[f"C{r}"] = sen; rs[f"C{r}"].font = F_BODY; rs[f"C{r}"].alignment = CTRV
    rs[f"D{r}"] = pres; rs[f"D{r}"].font = F_BODY; rs[f"D{r}"].alignment = TOP
    rs[f"E{r}"] = note; rs[f"E{r}"].font = F_BODY; rs[f"E{r}"].alignment = TOP
    for col in "ACDE":
        rs[f"{col}{r}"].border = HRULE
    rs.row_dimensions[r].height = 32

rs["A11"] = "Lato Laica S.p.A. (Offerta § 6 e § 9)"
rs["A11"].font = F_BODYB
rs["A12"] = ("Riccardo Dolcetta — Managing Director · Nicolò Zanuso — CFO · più, a seconda della sessione: "
             "CMO / Marketing Manager, CSO, referente servizio clienti, referenti IT e gestione dati, "
             "key user operativi.")
rs["A12"].font = F_BODY; rs["A12"].alignment = TOP
rs.merge_cells("A12:E12")
rs.row_dimensions[12].height = 30

# ================================================================ DATI FASI
# tipo, modalità, blocco, sett_inizio, durata_sett, n_incontri, durata_h, n_persone, gg_back
FASI = [
 dict(fase="Kick-off e analisi preliminare dei dati", tipo="Cliente", mod="Remoto",
      blocco="1 · Avvio", w0=1, wd=2, inc=1, ore=1.5, pers=3, back=5.0,
      descr="Incontro di avvio in remoto che sancisce l'apertura ufficiale del progetto: allineamento sulla "
            "governance, presentazione dei team e delle regole d'ingaggio tra tutti gli attori operativi e "
            "decisionali, condivisione della roadmap progettuale e del piano temporale di dettaglio con le "
            "date delle sessioni. Nella stessa sede si concordano le estrazioni da CRM ed ERP necessarie a "
            "comprendere la struttura dei dati di clienti e ordini, i punti di accesso all'informazione e la "
            "gestione dati attuale. L'analisi della base dati è condotta in modalità asincrona dai consulenti.",
      prep="Convocazione dell'incontro a cura del Commerciale; predisposizione delle slide di kick-off e "
           "della proposta di piano temporale (Gantt).",
      follow="E-mail di debrief con slide di kick-off e piano di lavoro definitivo allegati in PDF; analisi "
             "asincrona delle estrazioni ricevute.",
      inp="Estrazioni da CRM ed ERP su clienti e ordini; contatti dei referenti di progetto con adeguata "
          "autorità decisionale.",
      outp="Governance di progetto condivisa · piano temporale definitivo · e-mail di debrief con slide e "
           "piano in PDF · analisi preliminare della base dati.",
      owner="Commerciale (convocazione) → Project Manager (presidio)",
      part="Impresoft: 3 persone. Laica: Managing Director, CMO, CFO, CSO, referenti gestione dati / IT.",
      deliv="3 persone in remoto (1 facilitatore senior + 2 RevOps Consultant junior)",
      note="Offerta § 6.1 — 90 minuti, remoto. Il contratto indica 2 risorse Impresoft (PM + RevOps "
           "Consultant): qui sono 3 per scelta di pianificazione. L'analisi dati non è un momento erogato "
           "al cliente: è back office dentro questa fase. gg back office: STIMA, non desunta dal contratto."),

 dict(fase="Workshop #1 — Revenue model B2B e B2C", tipo="Cliente", mod="In presenza",
      blocco="2 · Workshop", w0=3, wd=1, inc=1, ore=6, pers=3, back=1.5,
      descr="Workshop in presenza per mappare le modalità con cui Laica genera valore sul canale B2B "
            "(distribuzione e OEM) e sui canali digitali B2C. Si analizza la proposta di valore per segmento, "
            "si mappano i canali di vendita e si esamina la struttura dei ricavi tra vendita dell'hardware e "
            "prodotti consumabili, costruendo il Revenue Model completo: value proposition, segmenti di "
            "clientela, competitor, canali, attività e risorse chiave, fonti di ricavo core, up-selling e "
            "cross-selling.",
      prep="Predisposizione del framework Revenue Model e condivisione dell'agenda.",
      follow="Compilazione del file di delivery, formalizzazione della mappa e restituzione al cliente; "
             "registrazione delle evidenze emerse.",
      inp="Disponibilità dei referenti di canale; dati di ricavo per linea di prodotto e per canale.",
      outp="Mappa formale del Revenue Model articolata per il canale B2B OEM e per i canali B2C.",
      owner="Project Manager",
      part="Impresoft: 3 persone. Laica: Managing Director, Marketing Manager, CFO, CSO.",
      deliv="3 persone in presenza (1 facilitatore senior + 2 RevOps Consultant junior)",
      note="Offerta § 6.2 — 6 ore, in presenza, UN SOLO incontro per entrambi i canali. gg back office: STIMA."),

 dict(fase="Workshop #2 — ICP, target group e customer journey", tipo="Cliente", mod="In presenza",
      blocco="2 · Workshop", w0=4, wd=2, inc=2, ore=6, pers=3, back=3.0,
      descr="Due sessioni in presenza — una dedicata al B2B/OEM e una al B2C/D2C — per definire i profili di "
            "cliente target e ricostruire il percorso d'acquisto nei mercati di riferimento. L'Ideal Customer "
            "Profile è costruito con metodologia SPICED; il customer journey è mappato lungo le fasi di "
            "acquisizione ed espansione del bowtie (consapevolezza, educazione, vendita, onboarding, "
            "soddisfazione, espansione), con la mappatura B2C estesa ai touchpoint proprietari e ai marketplace.",
      prep="Predisposizione delle schede ICP e della struttura di journey; condivisione dell'agenda delle "
           "due sessioni.",
      follow="Compilazione del file di delivery, formalizzazione delle schede e della mappa, restituzione al "
             "cliente; registrazione delle evidenze emerse.",
      inp="Disponibilità dei referenti marketing e sales per entrambe le sessioni.",
      outp="Schede ICP B2B e B2C per segmento · mappa dettagliata del Customer Journey multicanale.",
      owner="Project Manager",
      part="Impresoft: 3 persone. Laica: Managing Director, Marketing Manager, CFO, CSO.",
      deliv="3 persone in presenza × 2 sessioni (1 facilitatore senior + 2 RevOps Consultant junior)",
      note="Offerta § 6.3 — 12 ore complessive: 2 workshop dedicati (B2B + B2C) di circa 6 ore ciascuno, in "
           "presenza. È la fase più pesante del progetto. gg back office: STIMA."),

 dict(fase="Workshop #3 — Process design", tipo="Cliente", mod="In presenza",
      blocco="2 · Workshop", w0=6, wd=1, inc=1, ore=6, pers=3, back=2.0,
      descr="Workshop in presenza per mappare i flussi operativi correnti tra i dipartimenti: processi di "
            "Marketing, Sales, Service e Operations, con i relativi flussi e frizioni as-is. Si rileva cosa "
            "funziona e cosa non funziona nei processi in essere e si identificano i gap rispetto allo stato "
            "desiderato definito nel customer journey.",
      prep="Raccolta della documentazione di processo esistente; condivisione dell'agenda.",
      follow="Compilazione del file di delivery; avvio del diagramma dei flussi di processo as-is.",
      inp="Documentazione di processo; disponibilità dei process owner.",
      outp="Diagramma dei flussi di processo as-is — contributo del workshop alla gap analysis.",
      owner="Project Manager",
      part="Impresoft: 3 persone. Laica: Marketing Manager, CSO, referente servizio clienti, key user operativi.",
      deliv="3 persone in presenza (1 facilitatore senior + 2 RevOps Consultant junior)",
      note="Offerta § 6.4 — 6 ore, in presenza. Il perimetro contrattuale copre Marketing, Sales, Service e "
           "Operations. Nel contratto process design e user research sono un unico workshop: le interviste "
           "sono tenute su una riga a parte per separare i blocchi di progetto. gg back office: STIMA."),

 dict(fase="Interviste one-to-one — user research", tipo="Cliente", mod="Remoto",
      blocco="3 · User research", w0=7, wd=1, inc=5, ore=1, pers=1, back=1.0,
      descr="Cinque interviste individuali alle persone chiave di sales, marketing e customer care, condotte "
            "a valle del process design da un solo consulente. Sono interviste qualitative volte a raccogliere "
            "insight su come i processi sono vissuti quotidianamente e a far emergere le frizioni operative e "
            "tecnologiche che la mappatura di processo non rende visibili.",
      prep="Selezione degli intervistati con il cliente; predisposizione della guida di intervista.",
      follow="Sintesi trasversale delle interviste e integrazione nel documento di gap analysis.",
      inp="Individuazione e disponibilità delle 5 persone chiave da intervistare.",
      outp="Insight qualitativi · documento di sintesi dei punti di attrito operativi (gap analysis).",
      owner="Project Manager",
      part="Impresoft: 1 RevOps Consultant junior. Laica: 5 persone chiave di sales, marketing e customer care.",
      deliv="1 persona in remoto (1 RevOps Consultant junior)",
      note="Offerta § 6.4 — il contratto fissa 5 interviste one-to-one ma NON la loro durata: 1 ora ciascuna "
           "è una stima. Attività compresa nel Workshop #3, isolata qui come blocco a sé. gg back office: STIMA."),

 dict(fase="Workshop #4 — Tech stack e data governance", tipo="Cliente", mod="Remoto",
      blocco="2 · Workshop", w0=7, wd=1, inc=1, ore=4, pers=3, back=2.0,
      descr="Sessione in remoto per mappare le applicazioni in uso che impattano i processi di go-to-market: "
            "CRM HubSpot, gestionale/ERP, e-commerce, connettori marketplace. Si individua dove risiede "
            "l'informazione, come i dati fluiscono tra i sistemi e si analizzano le integrazioni esistenti, "
            "con l'obiettivo di fotografare lo stack as-is.",
      prep="Richiesta preventiva dell'inventario applicativo e delle licenze in essere.",
      follow="Compilazione del file di delivery; formalizzazione della mappa architetturale e degli schemi "
             "di integrazione.",
      inp="Inventario degli strumenti in uso; accesso ai referenti IT e, se necessario, ai fornitori software "
          "esterni.",
      outp="Mappa dell'architettura applicativa corrente · schema dei flussi di integrazione dati.",
      owner="Project Manager",
      part="Impresoft: 3 persone. Laica: referenti IT / sistemi informativi, key user di processo, eventuali "
           "fornitori software esterni.",
      deliv="3 persone in remoto (1 facilitatore senior + 2 RevOps Consultant junior)",
      note="Offerta § 6.5 — 4 ore, in remoto. gg back office: STIMA."),

 dict(fase="Workshop #5 — Data model e prioritizzazione", tipo="Cliente", mod="In presenza",
      blocco="2 · Workshop", w0=8, wd=1, inc=1, ore=6, pers=3, back=2.5,
      descr="Workshop in presenza per definire la struttura dei dati target — modello dati unificato per la "
            "gestione di cliente e prodotto — individuando i dati da tracciare lungo il customer journey, i "
            "KPI e i criteri di transizione. Nella seconda parte si consolidano le attività emerse e si "
            "prioritizzano per impatto, velocità a impatto e semplicità, classificandole con matrice di priorità.",
      prep="Predisposizione dell'albero dei KPI sul bowtie e della matrice di prioritizzazione.",
      follow="Formalizzazione del Data Model target; consolidamento delle priorità in vista della roadmap.",
      inp="Disponibilità dei dati necessari ad alimentare i KPI proposti.",
      outp="Data Model target · classificazione degli interventi per matrice di priorità.",
      owner="Project Manager",
      part="Impresoft: 3 persone. Laica: Managing Director, Marketing Manager, CFO, CSO, referente IT.",
      deliv="3 persone in presenza (1 facilitatore senior + 2 RevOps Consultant junior)",
      note="Offerta § 6.6 — 6 ore, in presenza. Il contratto indica 4 risorse Impresoft sul § 6.6 (RevOps "
           "Consultant, CX Specialist, Integration Expert, PM): qui sono 3. La roadmap per wave non è una "
           "fase: è contenuto del report finale. gg back office: STIMA."),

 dict(fase="Review con il cliente", tipo="Cliente", mod="Remoto",
      blocco="4 · Sintesi e validazione", w0=9, wd=1, inc=1, ore=None, pers=3, back=0.5,
      descr="Incontro di validazione con il team di lavoro lato cliente sui business case e sulle priorità "
            "consolidate, prima della formalizzazione del report. Si verifica l'allineamento su priorità e "
            "numeri e si stabilisce se procedere o iterare su specifici ambiti.",
      prep="Invio anticipato dei business case e della matrice di priorità per lettura preventiva.",
      follow="Recepimento delle iterazioni richieste e aggiornamento del report in costruzione.",
      inp="Presenza dei decisori lato cliente.",
      outp="Business case e priorità validati, oppure lista delle iterazioni richieste.",
      owner="Project Manager",
      part="Impresoft: 3 persone. Laica: team di lavoro e decisori.",
      deliv="3 persone in remoto (1 facilitatore senior + 2 RevOps Consultant junior)",
      note="FUORI PERIMETRO CONTRATTUALE — l'Offerta non prevede una sessione di review tra il Workshop #5 e "
           "il meeting di presentazione. Mantenuta su richiesta. DURATA DA DEFINIRE: finché la cella è vuota "
           "le giornate front di questa fase valgono 0. gg back office: STIMA."),

 dict(fase="Consolidamento report, roadmap e business case", tipo="Interna", mod="Interna",
      blocco="4 · Sintesi e validazione", w0=9, wd=2, inc=None, ore=None, pers=None, back=7.5,
      descr="Fase interna di consolidamento dell'intero assessment: costruzione del report consolidato, della "
            "roadmap strategica di implementazione suddivisa per wave secondo le priorità individuate e del "
            "business case economico analitico a supporto. Il documento viene anticipato al cliente via e-mail "
            "in PDF prima del meeting di presentazione.",
      prep="Consolidamento dei file di delivery di tutte le sessioni precedenti e degli esiti della review.",
      follow="Invio anticipato del PDF al cliente in vista del meeting di presentazione.",
      inp=None,
      outp="Report consolidato di assessment · roadmap per wave · business case economico · PDF in lingua "
           "italiana.",
      owner="Project Manager",
      part="Team di progetto",
      deliv=None,
      note="Offerta § 6.6 e § 11 — l'invio anticipato del PDF è la milestone che attiva la fatturazione del "
           "50% residuo. gg back office: STIMA, ed è la voce più pesante del progetto."),

 dict(fase="Meeting di presentazione dell'output", tipo="Cliente", mod="In presenza",
      blocco="5 · Chiusura", w0=10, wd=1, inc=1, ore=2, pers=3, back=1.0,
      descr="Presentazione in presenza dell'output finale dell'assessment — report consolidato, roadmap per "
            "wave e business case — come passaggio verso le successive scelte di investimento tecnologico e "
            "organizzativo. Si svolge circa 15 giorni dopo il workshop Data Model.",
      prep="Predisposizione delle slide di presentazione; conferma di agenda e partecipanti.",
      follow="Raccolta delle precisazioni e micro-rettifiche emerse in incontro, da recepire nella versione "
             "inglese del report.",
      inp="Presenza dei decisori lato cliente.",
      outp="Assessment presentato e consegnato · next step condivisi.",
      owner="Project Manager",
      part="Impresoft: 3 persone. Laica: Managing Director, Marketing Manager, CFO, CSO, referente IT.",
      deliv="3 persone in presenza (1 facilitatore senior + 2 RevOps Consultant junior)",
      note="Offerta § 6.6 — 2 ore, in presenza, circa 15 giorni dopo il Workshop #5. VINCOLO DI "
           "PIANIFICAZIONE da rispettare nella colonna «Sett. inizio». gg back office: STIMA."),

 dict(fase="Redazione output in lingua inglese", tipo="Interna", mod="Interna",
      blocco="5 · Chiusura", w0=11, wd=1, inc=None, ore=None, pers=None, back=2.0,
      descr="Redazione della versione inglese del report di assessment, che il contratto colloca "
            "esplicitamente dopo il meeting di presentazione per poter includere le precisazioni e le "
            "micro-rettifiche emerse durante l'incontro.",
      prep="Recepimento delle rettifiche emerse in presentazione.",
      follow="Consegna del PDF in lingua inglese.",
      inp=None,
      outp="Report di assessment in lingua inglese (PDF).",
      owner="Project Manager",
      part="Team di progetto",
      deliv=None,
      note="Offerta § 6.6 — deliverable obbligatorio. ATTENZIONE: il § 6.7 esclude dall'offerta la "
           "«traduzione di testi» come servizio di terze parti, quindi la versione inglese è a carico "
           "interno. Questa coda cade OLTRE le 10 settimane dichiarate al § 8. gg back office: STIMA."),
]

# ================================================================ WBS
ws = wb.create_sheet("WBS", 0)
ws.sheet_view.showGridLines = False

COLS = [
    ("A", "#", 5, "num"), ("B", "Cod.\nWBS", 8, "txt"), ("C", "Fase", 34, "txt"),
    ("D", "Tipo", 11, "txt"), ("E", "Modalità", 13, "txt"), ("F", "Blocco", 22, "txt"),
    ("G", "Sett.\ninizio", 8, "in"), ("H", "Durata\n(sett.)", 8, "in"), ("I", "Sett.\nfine", 8, "calc"),
    ("J", "Descrizione", 70, "txt"), ("K", "Attività preparatorie", 40, "txt"),
    ("L", "Attività di follow-up", 40, "txt"), ("M", "Input dal cliente", 34, "txt"),
    ("N", "Output / Deliverable", 44, "txt"), ("O", "Owner", 26, "txt"), ("P", "Partecipanti", 34, "txt"),
    ("Q", "N.\nincontri", 8, "in"), ("R", "Durata\nincontro (h)", 10, "in"),
    ("S", "N.\npersone", 8, "in"),
    ("T", "gg Front\noffice", 10, "calc"), ("U", "gg Back\noffice", 10, "in"),
    ("V", "gg\nViaggio", 9, "calc"), ("W", "Delivery (composizione team)", 40, "txt"),
    ("X", "Tot.\ngiornate", 10, "calc"), ("Y", "Ore\ntotali", 10, "calc"),
    ("Z", "Costo totale", 15, "calc"), ("AA", "Note", 52, "txt"),
]
LAST = "AA"
HR, R0 = 5, 6
R1 = R0 + len(FASI) - 1
RT = R1 + 1

banda_titolo(ws, 1, "A", LAST,
             "WORK BREAKDOWN STRUCTURE — REVOPS ASSESSMENT B2B E B2C · LAICA S.P.A.")

ws.merge_cells(f"A2:{LAST}2")
ws["A2"] = ("Allineata all'Offerta CX260803_Off_Laica_RevOps assessment B2B e B2C v1.1 del 28 ago 2026.  ·  "
            "LEGENDA — Colonne con intestazione GIALLA: da compilare a mano, celle con riquadro nero. "
            "Celle su fondo GRIGIO: calcolate da formula. Tariffe, ore per giornata e costi di trasferta si "
            "impostano nel foglio «Parametri»; i nomi del team nel foglio «Ruoli».")
ws["A2"].font = F_SUB; ws["A2"].alignment = TOP
ws.row_dimensions[2].height = 26

banda(ws, 3, "A", LAST,
      "gg Front office e gg Viaggio sono CALCOLATE: front = n. incontri × durata × n. persone ÷ ore per "
      "giornata; viaggio = n. incontri × n. persone × ore di viaggio ÷ ore per giornata, solo sulle fasi in "
      "presenza. L'unico input di effort è gg Back office — e il contratto non lo quantifica in nessun punto: "
      "quei valori sono STIME Impresoft.",
      FILL_YEL, F_ACC, 30)
ws["A3"].alignment = TOP

for letter, header, width, kind in COLS:
    ws.column_dimensions[letter].width = width
    c = ws[f"{letter}{HR}"]
    c.value = header
    c.fill = FILL_YEL if kind == "in" else FILL_BLK
    c.font = F_HEADY if kind == "in" else F_HEAD
    c.alignment = CTR
ws.row_dimensions[HR].height = 34
ws[f"T{HR}"].comment = Comment(
    "Calcolata: N. incontri × Durata incontro (h) × N. persone ÷ Ore per giornata.\n"
    "Per cambiare l'impegno, agire su quelle tre colonne.", "WBS")
ws[f"U{HR}"].comment = Comment(
    "Stime Impresoft. L'Offerta quantifica solo le ore erogate al cliente e il numero "
    "di trasferte: sul back office non dice nulla.", "WBS")

for i, f in enumerate(FASI):
    r = R0 + i
    vals = {
        "A": i + 1, "B": f"1.{i+1}",
        "C": f["fase"], "D": f["tipo"], "E": f["mod"], "F": f["blocco"],
        "G": f["w0"], "H": f["wd"],
        "I": f'=IF(AND(G{r}<>"",H{r}<>""),G{r}+H{r}-1,"")',
        "J": f["descr"], "K": f["prep"], "L": f["follow"], "M": f["inp"], "N": f["outp"],
        "O": f["owner"], "P": f["part"],
        "Q": f["inc"], "R": f["ore"], "S": f["pers"],
        "T": f'=IF(OR(Q{r}="",R{r}="",S{r}=""),0,Q{r}*R{r}*S{r}/{P_ORE})',
        "U": f["back"],
        "V": f'=IF(AND(E{r}="In presenza",Q{r}<>"",S{r}<>""),Q{r}*S{r}*{P_HVIAG}/{P_ORE},0)',
        "W": f["deliv"],
        "X": f"=T{r}+U{r}+V{r}",
        "Y": f"=X{r}*{P_ORE}",
        "Z": f"=T{r}*{P_ORE}*{P_FRO}+U{r}*{P_ORE}*{P_BACK}+V{r}*{P_ORE}*{P_TRAV}",
        "AA": f["note"],
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
            c.number_format = EUR if letter == "Z" else NUM
        elif kind == "num":
            c.font = F_BODYB; c.alignment = CTRV; c.border = HRULE
        else:
            c.font = F_BODYB if letter == "C" else F_BODY
            c.alignment = TOP; c.border = HRULE
    for letter in ("G", "H", "I", "Q", "S"):
        ws[f"{letter}{r}"].number_format = INT
    ws.row_dimensions[r].height = 96

banda(ws, RT, "A", "P", "TOTALE PROGETTO", FILL_BLK, F_TOTAL, 22, OVERY)
ws[f"A{RT}"].alignment = RGTC
for letter, _, _, _ in COLS:
    c = ws[f"{letter}{RT}"]
    c.fill = FILL_BLK; c.font = F_TOTAL; c.alignment = CTR; c.border = OVERY
for letter in ("T", "U", "V", "X", "Y", "Z"):
    ws[f"{letter}{RT}"] = f"=SUM({letter}{R0}:{letter}{R1})"
    ws[f"{letter}{RT}"].number_format = EUR if letter == "Z" else NUM
ws[f"Q{RT}"] = f"=SUM(Q{R0}:Q{R1})"
ws[f"Q{RT}"].number_format = INT

dv_tipo = DataValidation(type="list", formula1='"Cliente,Interna"', allow_blank=True)
dv_mod  = DataValidation(type="list", formula1='"In presenza,Remoto,Interna"', allow_blank=True)
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

banda_titolo(ts, 1, "A", W_LAST, "TIMELINE DI PROGETTO — 10 SETTIMANE + CODA DI TRADUZIONE")
ts.merge_cells(f"A2:{W_LAST}2")
ts["A2"] = ("Le barre si disegnano da sole dalle settimane impostate nella WBS: qui non si scrive nulla. "
            "GIALLO = i 5 workshop, il nucleo erogato dell'assessment. NERO = altra sessione con il cliente. "
            "GRIGIO = fase interna. La riga in fondo conta quante fasi sono attive in ciascuna settimana. "
            "Vincolo contrattuale: il meeting di presentazione cade circa 15 giorni dopo il Workshop #5.")
ts["A2"].font = F_SUB; ts["A2"].alignment = TOP
ts.row_dimensions[2].height = 30

for letter, txt, w in (("A", "#", 5), ("B", "Fase", 38), ("C", "Tipo", 11),
                       ("D", "Blocco", 22), ("E", "Da\nsett.", 8), ("F", "A\nsett.", 8)):
    ts.column_dimensions[letter].width = w
    c = ts[f"{letter}{THR}"]
    c.value = txt; c.font = F_HEAD; c.fill = FILL_BLK; c.alignment = CTR
for k in range(NW):
    letter = get_column_letter(W_FIRST + k)
    ts.column_dimensions[letter].width = 5.5
    c = ts[f"{letter}{THR}"]
    c.value = k + 1; c.number_format = '"S"0'
    c.font = F_HEAD; c.fill = FILL_BLK; c.alignment = CTR
ts.row_dimensions[THR].height = 24

banda(ts, RMK, "A", "F", "Settimane con fasi in parallelo  ›", FILL_WHT, F_BODYB, 18)
ts[f"A{RMK}"].alignment = RGTC
for k in range(NW):
    letter = get_column_letter(W_FIRST + k)
    c = ts[f"{letter}{RMK}"]
    c.value = f'=IF({letter}{RA}>1,"›","")'
    c.font = F_ACC; c.alignment = CTR
ts.conditional_formatting.add(f"{GC}{RMK}:{W_LAST}{RMK}", FormulaRule(
    formula=[f'{GC}{RMK}<>""'], fill=FILL_YEL,
    font=Font(name=FONT, size=9, bold=True, color=BLK), stopIfTrue=True))

for i in range(len(FASI)):
    r, wr = T0 + i, R0 + i
    for letter, src in (("A", "A"), ("B", "C"), ("C", "D"), ("D", "F"), ("E", "G"), ("F", "I")):
        c = ts[f"{letter}{r}"]
        c.value = f"=WBS!{src}{wr}"
        c.font = F_CALCB if letter == "B" else F_CALC
        c.border = HRULE
        c.alignment = TOP if letter in ("B", "D") else CTRV
        if letter in ("A", "E", "F"):
            c.number_format = INT
    for k in range(NW):
        letter = get_column_letter(W_FIRST + k)
        c = ts[f"{letter}{r}"]
        c.value = (f'=IF(AND($E{r}<>"",$F{r}<>"",$E{r}<={letter}${THR},'
                   f'$F{r}>={letter}${THR}),1,"")')
        c.font = F_BODY
        c.number_format = ";;;"
        c.border = HRULE
        c.alignment = CTR
    ts.row_dimensions[r].height = 24

GRID = f"{GC}{T0}:{W_LAST}{T1}"
ts.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'AND(ISNUMBER(SEARCH("Workshop",$D{T0})),{GC}{T0}=1)'],
    fill=FILL_YEL, stopIfTrue=True))
ts.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'AND($C{T0}="Interna",{GC}{T0}=1)'], fill=FILL_GMID, stopIfTrue=True))
ts.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'{GC}{T0}=1'], fill=FILL_BLK, stopIfTrue=True))
ts.conditional_formatting.add(f"B{T0}:F{T1}", FormulaRule(
    formula=[f'ISNUMBER(SEARCH("Workshop",$D{T0}))'],
    font=Font(name=FONT, size=9, bold=True, color=BLK)))

banda(ts, RA, "A", "F", "Fasi attive nella settimana", FILL_BLK, F_TOTAL, 22, OVERY)
ts[f"A{RA}"].alignment = RGTC
for k in range(NW):
    letter = get_column_letter(W_FIRST + k)
    c = ts[f"{letter}{RA}"]
    c.value = f"=COUNT({letter}{T0}:{letter}{T1})"
    c.font = F_TOTAL; c.fill = FILL_BLK; c.alignment = CTR
    c.border = OVERY; c.number_format = INT
ts.conditional_formatting.add(f"{GC}{RA}:{W_LAST}{RA}", FormulaRule(
    formula=[f'{GC}{RA}>1'], fill=FILL_YEL,
    font=Font(name=FONT, size=9, bold=True, color=BLK), stopIfTrue=True))

RB = RA + 2
for r, label, formula in (
    (RB,     "Sessioni erogate al cliente (n. incontri)", f"=WBS!Q{RT}"),
    (RB + 1, "di cui in presenza — trasferte incluse nell'investimento",
     f'=SUMPRODUCT((WBS!$E${R0}:$E${R1}="In presenza")*WBS!$Q${R0}:$Q${R1})'),
    (RB + 2, "Settimana di massima concentrazione (n. fasi in parallelo)",
     f"=MAX({GC}{RA}:{W_LAST}{RA})"),
):
    ts[f"A{r}"] = label; ts[f"A{r}"].font = F_BODYB
    ts.merge_cells(f"A{r}:E{r}")
    c = ts[f"F{r}"]
    c.value = formula
    c.font = F_CALCB; c.fill = FILL_S2; c.border = CALCBOX
    c.alignment = CTR; c.number_format = INT
ts.freeze_panes = f"{GC}{T0}"

# ================================================================ RIEPILOGO
sm = wb.create_sheet("Riepilogo", 2)
sm.sheet_view.showGridLines = False
for col, w in (("A", 60), ("B", 2), ("C", 18), ("D", 15), ("E", 64)):
    sm.column_dimensions[col].width = w

banda_titolo(sm, 1, "A", "E", "RIEPILOGO DI PROGETTO")
sm.merge_cells("A2:E2")
sm["A2"] = ("Tutti i valori sono calcolati dalla WBS e dai Parametri. Nessuna cella di questo foglio va "
            "compilata a mano.")
sm["A2"].font = F_SUB
sm.row_dimensions[2].height = 18


def sect(r, title):
    banda(sm, r, "A", "E", f"›  {title}", FILL_YEL,
          Font(name=FONT, size=10, bold=True, color=BLK), 20)


def line(r, label, formula, fmt=NUM, unit="", note="", bold=False):
    sm[f"A{r}"] = label
    sm[f"A{r}"].font = F_BODYB if bold else F_BODY
    sm[f"A{r}"].alignment = TOP; sm[f"A{r}"].border = HRULE
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
E_RNG = f"WBS!$E${R0}:$E${R1}"
Q_RNG = f"WBS!$Q${R0}:$Q${R1}"
S_RNG = f"WBS!$S${R0}:$S${R1}"
X_RNG = f"WBS!$X${R0}:$X${R1}"
Z_RNG = f"WBS!$Z${R0}:$Z${R1}"

sect(4, "IMPEGNO COMPLESSIVO")
line(5, "Giornate front office (tempo erogato al cliente)", f"=WBS!T{RT}", NUM, "gg-uomo",
     "Calcolate da n. incontri × durata × n. persone. Incontri in remoto inclusi.")
line(6, "Giornate back office", f"=WBS!U{RT}", NUM, "gg-uomo",
     "STIME Impresoft: il contratto non quantifica il back office.")
line(7, "Giornate di viaggio", f"=WBS!V{RT}", NUM, "gg-uomo",
     "Solo fasi in presenza. Incluse nel monte ore.")
line(8, "TOTALE giornate-uomo di progetto", f"=WBS!X{RT}", NUM, "gg-uomo", "", bold=True)
line(9, "TOTALE ore di progetto", f"=WBS!Y{RT}", NUM, "h",
     f"Totale giornate × ore per giornata ({P_ORE}).", bold=True)

sect(11, "COSTO DELLE GIORNATE")
line(12, "Costo ore front office", f"=WBS!T{RT}*{P_ORE}*{P_FRO}", EUR, "€")
line(13, "Costo ore back office", f"=WBS!U{RT}*{P_ORE}*{P_BACK}", EUR, "€")
line(14, "Costo ore di viaggio", f"=WBS!V{RT}*{P_ORE}*{P_TRAV}", EUR, "€")
line(15, "COSTO TOTALE DELLE GIORNATE", f"=WBS!Z{RT}", EUR, "€", "", bold=True)
line(16, "Quadratura (deve coincidere con la riga sopra)", "=SUM(C12:C14)", EUR, "€",
     "Controllo interno: se differisce c'è un errore di formula.")

sect(18, "COSTO DELLE TRASFERTE — già incluso nell'investimento (Offerta § 10.1)")
line(19, "Sessioni in presenza",
     f'=SUMPRODUCT(({E_RNG}="In presenza")*{Q_RNG})', INT, "n.",
     "Dalla WBS. L'Offerta § 10.1 include le trasferte per 6 incontri in presenza.")
line(20, "Persone-sessione in trasferta",
     f'=SUMPRODUCT(({E_RNG}="In presenza")*{Q_RNG}*{S_RNG})', INT, "n.",
     "Somma delle presenze fisiche: determina il vitto.")
line(21, "Viaggi in auto", f"=IF({P_PAUTO}=0,0,ROUNDUP(C20/{P_PAUTO},0))", INT, "n.",
     "Persone-sessione ÷ persone per auto. Determina km e pedaggi.")
line(22, "Ammortamento ACI", f"=C21*{P_KM}*{P_ACI}", EUR, "€",
     f"Viaggi × km a/r ({P_KM}) × tariffa €/km ({P_ACI}).")
line(23, "Pedaggi — casello", f"=C21*{P_PED}", EUR, "€")
line(24, "Vitto", f"=C20*{P_VITTO}", EUR, "€", "Persone-sessione × costo per persona.")
line(25, "TOTALE COSTO TRASFERTE", "=SUM(C22:C24)", EUR, "€", "", bold=True)

sect(27, "COSTO PIENO E MARGINE")
line(28, "Costo delle giornate", "=C15", EUR, "€")
line(29, "Costo delle trasferte", "=C25", EUR, "€")
line(30, "COSTO PIENO DI PROGETTO", "=C28+C29", EUR, "€",
     "Giornate + trasferte. È il costo da confrontare con l'investimento.", bold=True)
line(31, "Investimento contrattuale (una tantum)", f"={P_INV}", EUR, "€",
     "Offerta § 10. Le trasferte dei 6 incontri in presenza sono già dentro questa cifra.")
line(32, "MARGINE DI PROGETTO", "=C31-C30", EUR, "€", "", bold=True)
line(33, "Margine percentuale", '=IF(C31=0,"",C32/C31)', "0.0%", "%",
     "Se negativo la cella si inverte in nero.")
line(34, "Tariffa media implicita sulle ore erogate",
     '=IF(C9=0,"",C31/C9)', EUR, "€/h",
     "Investimento ÷ totale ore. Da confrontare con le tre tariffe orarie impostate nei Parametri.",
     bold=True)
for cella in ("C32", "C33"):
    sm.conditional_formatting.add(cella, FormulaRule(
        formula=[f"{cella}<0"], fill=FILL_BLK,
        font=Font(name=FONT, size=9, bold=True, color=WHT), stopIfTrue=True))

sect(36, "RIPARTIZIONE PER TIPO DI FASE")
line(37, "Giornate — fasi con il cliente", f'=SUMIF({D_RNG},"Cliente",{X_RNG})', NUM, "gg-uomo")
line(38, "Giornate — fasi interne", f'=SUMIF({D_RNG},"Interna",{X_RNG})', NUM, "gg-uomo")
line(39, "Costo — fasi con il cliente", f'=SUMIF({D_RNG},"Cliente",{Z_RNG})', EUR, "€")
line(40, "Costo — fasi interne", f'=SUMIF({D_RNG},"Interna",{Z_RNG})', EUR, "€")
line(41, "Fasi con il cliente / fasi interne",
     f'=COUNTIF({D_RNG},"Cliente")&" / "&COUNTIF({D_RNG},"Interna")', "General", "")

sect(43, "LEAD TIME")
line(44, "Somma delle durate delle fasi, se svolte in sequenza",
     f"=SUM(WBS!H{R0}:H{R1})", INT, "settimane")
line(45, "Lead time di calendario effettivo",
     f"=MAX(WBS!I{R0}:I{R1})-MIN(WBS!G{R0}:G{R1})+1", INT, "settimane",
     "Prima settimana di inizio → ultima settimana di fine, fasi sovrapposte incluse.", bold=True)
line(46, "Settimane recuperate dalla parallelizzazione", "=C44-C45", INT, "settimane")
line(47, "Lead time target dichiarato", f"={P_TGT}", INT, "settimane", "Offerta § 8.")
line(48, "SCOSTAMENTO effettivo vs target", "=C45-C47", INT, "settimane",
     "Se diverso da zero la cella si inverte in nero.", bold=True)
sm.conditional_formatting.add("C48", FormulaRule(
    formula=["C48<>0"], fill=FILL_BLK,
    font=Font(name=FONT, size=9, bold=True, color=WHT), stopIfTrue=True))
sm["A49"] = "Esito del controllo"
sm["A49"].font = F_BODYB; sm["A49"].border = HRULE
sm.merge_cells("C49:E49")
sm["C49"] = ('=IF(C45=C47,"Il piano quadra con il lead time target.",'
             '"Scostamento sul target: la coda di redazione in lingua inglese cade oltre le 10 settimane.")')
sm["C49"].font = F_BODY; sm["C49"].alignment = LFTC
sm.conditional_formatting.add("C49:E49", FormulaRule(
    formula=["$C$45<>$C$47"], fill=FILL_BLK,
    font=Font(name=FONT, size=9, bold=True, color=WHT), stopIfTrue=True))
sm.row_dimensions[49].height = 18

sect(51, "IPOTESI, SCOSTAMENTI E PUNTI APERTI")
NOTES = [
 ("Fonte del piano",
  "Offerta CX260803_Off_Laica_RevOps assessment B2B e B2C v1.1 del 28 ago 2026. Modalità, ore e numero di "
  "sessioni di ciascuna fase sono letti dal capitolo 6 del contratto. Le 6 sessioni in presenza coincidono "
  "con le 6 trasferte incluse nell'investimento al § 10.1."),
 ("Risorse per sessione: 3 anziché 2",
  "Il contratto indica 2 risorse Impresoft su kick-off e Workshop #1-#4, e 4 sul § 6.6 (Workshop #5 e "
  "presentazione). La WBS pianifica 3 persone su tutte le sessioni con il cliente — 1 facilitatore senior + "
  "2 RevOps Consultant junior — come da indicazione. Sul Workshop #2, che vale 2 sessioni da 6 ore, la terza "
  "persona costa 1,5 giornate front in più."),
 ("Ruoli generici invece di ruoli specialistici",
  "Il contratto nomina CX Strategist, Process Analyst e Integration Expert su workshop specifici. Nella WBS "
  "le persone sono indicate come «RevOps Consultant junior» senza specializzazione funzionale, per scelta di "
  "pianificazione. Sulle sessioni corrispondenti il contratto promette al cliente una competenza dedicata."),
 ("Giornate back office: stime, non contratto",
  "Il contratto quantifica solo le ore erogate al cliente e il numero di trasferte. Le giornate back office "
  "della WBS sono stime Impresoft e sono la variabile su cui si gioca il margine: la voce più pesante è il "
  "consolidamento del report, roadmap e business case."),
 ("Interviste one-to-one",
  "Il contratto (§ 6.4) le colloca DENTRO il Workshop #3 e ne fissa il numero — 5 — ma non la durata. "
  "Nella WBS sono un blocco a sé per separare i momenti di progetto; 1 ora ciascuna è una stima."),
 ("Review con il cliente: fuori perimetro",
  "L'Offerta non prevede una sessione di validazione tra il Workshop #5 e il meeting di presentazione. La "
  "fase 8 è mantenuta su richiesta: la sua durata è da definire e finché è vuota non genera giornate front."),
 ("Roadmap",
  "Non è una fase. È contenuto del Workshop #5 (prioritizzazione) e deliverable del report finale, come "
  "previsto dal § 6.6."),
 ("Vincolo dei 15 giorni",
  "Il § 6.6 fissa il meeting di presentazione a circa 15 giorni dopo il Workshop #5. Nel piano proposto: "
  "Workshop #5 in S8, presentazione in S10. Spostando una delle due settimane, il vincolo va riverificato."),
 ("Output in lingua inglese",
  "Deliverable obbligatorio del § 6.6, da redigere DOPO il meeting di presentazione. Ma il § 6.7 esclude "
  "dall'offerta la «traduzione di testi» come servizio di terze parti: la versione inglese è a carico "
  "interno. È la fase 11 e cade oltre le 10 settimane dichiarate al § 8."),
 ("Durate contrattuali dichiarate due volte",
  "Il § 6 dichiara «8-10 settimane», il § 8 dichiara «10 settimane». Incoerenza interna al contratto, da "
  "sanare prima della firma."),
 ("Fatturazione",
  "50% alla sottoscrizione, 50% nel mese in cui il PDF con l'output viene anticipato via e-mail al cliente "
  "(§ 11). Il trigger della seconda fattura è la fase 9, non la presentazione."),
 ("Identità visiva",
  "Palette Impresoft di tre colori (giallo #FDC300, nero, bianco) più grigi puri derivati dal nero. Font "
  "Manrope, fallback Arial. Nessun colore semantico: gli stati usano la scala monocroma e le parole."),
]
r = 52
for titolo, testo in NOTES:
    sm[f"A{r}"] = titolo
    sm[f"A{r}"].font = F_BODYB; sm[f"A{r}"].alignment = TOP; sm[f"A{r}"].border = HRULE
    sm.merge_cells(f"C{r}:E{r}")
    sm[f"C{r}"] = testo
    sm[f"C{r}"].font = F_BODY; sm[f"C{r}"].alignment = TOP
    sm.row_dimensions[r].height = 46
    r += 1

# ================================================================ RIFINITURE
for sh, land in ((ws, True), (ts, True), (sm, False), (ps, False), (rs, False)):
    sh.page_setup.orientation = "landscape" if land else "portrait"
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr.fitToPage = True
    sh.print_title_rows = f"{HR}:{HR}" if sh is ws else None

wb.save(OUT)
print("scritto:", OUT)
