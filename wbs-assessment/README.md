# WBS — Progetto di Assessment Go-to-Market

`WBS_Assessment_GTM.xlsx` è il file di lavoro da editare al desktop.
`build_wbs.py` lo rigenera da zero (`python3 build_wbs.py`) — attenzione: sovrascrive
il file, quindi non usarlo dopo aver iniziato a compilare a mano.

## Fogli

| Foglio | Contenuto |
|---|---|
| **WBS** | Le 12 fasi, una per riga. Colonne con intestazione gialla = da compilare, celle grigie = calcolate |
| **Timeline** | Gantt per settimana, si disegna da sé dalle settimane impostate nella WBS |
| **Riepilogo** | Totali, ripartizioni, controllo del lead time. Nessuna cella da compilare |
| **Parametri** | Ore per giornata e le tre tariffe orarie (front / back / viaggio) |
| **Ruoli** | Mappatura ruolo → nome del team di progetto |

## Da compilare

1. Tariffe orarie front / back / viaggio nel foglio **Parametri** (finché sono vuote, i costi restano a zero)
2. Giornate front / back / viaggio per ciascuna fase nel foglio **WBS**
3. Durata degli incontri e modalità delle fasi marcate «Da definire»
4. Nomi del team nel foglio **Ruoli**

## Convenzione sulle giornate

Le colonne giornate esprimono **giornate-uomo aggregate**, non giorni di calendario:
2 consulenti per mezza giornata dal cliente = **1** giornata front.
La composizione della squadra si descrive nella colonna «Delivery».

## Identità visiva

Il file segue la skill `impresoft-brand` (Impresoft Brand Manual & Corporate
Guidelines, agg. 27.04.2023) e la ricetta Excel di `references/office-docs.md`.

- **Palette di tre colori**: giallo `#FDC300`, nero `#000000`, bianco `#FFFFFF`,
  più grigi puri derivati dal nero (`#F4F4F4` `#3A3A3A` `#6E6E6E` `#8C8C8C`
  `#9B9B9B` `#DCDCDC`). Nessun altro colore compare nel file.
- **Font**: Manrope. Se non è installato sulle macchine che apriranno il file,
  cambiare la costante `FONT` in cima a `build_wbs.py` con `"Arial"` (fallback
  indicato dal manuale) e rigenerare.
- **Codifica visiva**: il giallo marca le intestazioni delle colonne da
  compilare e, nella Timeline, le fasi che si svolgono in parallelo — la serie
  che porta il messaggio. Le altre fasi scalano in monocromia (nero = con il
  cliente, grigio = interna). Le celle da compilare portano un riquadro nero.
- **Nessun colore semantico**: il brand non prevede verde «ok» e rosso
  «errore». Gli stati sono codificati con la scala monocroma (fondo nero,
  testo bianco) e con etichette in parole — vedi «Esito del controllo» nel
  Riepilogo.

### Nota tecnica su openpyxl

I fill vanno dichiarati in ARGB a 8 cifre con `start_color` **e** `end_color`
(vedi la funzione `fill()` in `build_wbs.py`). openpyxl omette l'attributo
colore quando il valore coincide con il proprio default, e `000000`
normalizzato è esattamente quel default: un fill nero dichiarato in modo
abbreviato sparisce dai `dxf` della formattazione condizionale, e le barre
nere del Gantt non vengono disegnate.

### Nota sul ricalcolo

Il file consegnato **non** passa da LibreOffice: il ricalcolo viene eseguito su
una copia usa e getta, perché la riscrittura di LibreOffice inietta nei fogli
di stile record fuori palette (`#003300`, `#FF9900`) e sostituzioni di font.
Le formule sono verificate su quella copia; il file consegnato ricalcola
all'apertura in Excel.
