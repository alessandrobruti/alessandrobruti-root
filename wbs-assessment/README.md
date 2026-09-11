# WBS — RevOps Assessment B2B e B2C · Laica S.p.A.

`260803_LAICA_WBS_RevOps_Assessment_V.1.1.xlsx` è il file di lavoro da editare al desktop.
`build_wbs.py` lo rigenera da zero (`python3 build_wbs.py`) — sovrascrive il file,
quindi non usarlo dopo aver iniziato a compilare a mano.

**Fonte:** Offerta `CX260803_Off_Laica_RevOps assessment B2B e B2C` v1.1 del 28 ago 2026.
Modalità, ore e numero di sessioni di ogni fase sono letti dal capitolo 6 del contratto.

## Fogli

| Foglio | Contenuto |
|---|---|
| **WBS** | Le 11 fasi, una per riga. Intestazione gialla = da compilare, celle grigie = calcolate |
| **Timeline** | Gantt per settimana, si disegna da sé dalle settimane impostate nella WBS. In alto la settimana ISO di partenza: unico comando per spostare il piano sul calendario |
| **Riepilogo** | Impegno, costo giornate, costo trasferte, costo pieno e margine, lead time, punti aperti |
| **Parametri** | Ore/giornata, tre tariffe orarie, ore di viaggio, investimento, costi di trasferta |
| **Ruoli** | Mappatura ruolo → nome del team di progetto |

## Settimane relative e shift sul calendario

La WBS ragiona **sempre in settimane relative**: il progetto parte dalla settimana 1
e il lead time si misura da lì. La settimana di calendario reale si imposta in **una
sola cella**, in alto nel foglio Timeline (`F2`, gialla): la settimana 1 diventa quella,
e le etichette di colonna e la fascia dei mesi si spostano di conseguenza. Le barre non
si muovono, perché dipendono dalle settimane relative della WBS.

Serve a fare proiezioni: settimana 41 → il piano chiude alla 50; settimana 50 → chiude
alla 59, con la fascia dei mesi che rotola correttamente su *Gennaio 2027*. L'anno di
riferimento sta nei Parametri e serve a ricavare i mesi.

Il mese si ricava dal lunedì della settimana ISO:
`DATE(anno,1,4) - WEEKDAY(DATE(anno,1,4),3) + 7*(settimana-1)` — il 4 gennaio cade
sempre nella settimana ISO 1, e `WEEKDAY(...,3)` vale 0 di lunedì. La formula regge i
numeri oltre la 52: rotola nell'anno successivo, che è il comportamento voluto.

La fascia dei mesi usa le **sigle di tre lettere** su ogni settimana, non il nome intero
centrato sul gruppo: una cella con formula che restituisce `""` non è vuota per Excel e
blocca lo sbordo del testo, quindi un nome lungo resterebbe tagliato. I nomi per intero,
con gli anni, stanno nella riga «Periodo di calendario coperto».

## Come si calcola l'impegno

Le giornate front office e viaggio **non si compilano**, si calcolano:

```
gg Front office = N. incontri × Durata incontro (h) × N. persone ÷ Ore per giornata
gg Viaggio      = N. incontri × N. persone × Ore di viaggio ÷ Ore per giornata
                  (solo sulle fasi in presenza)
```

Per cambiare l'impegno si agisce su **N. incontri**, **Durata incontro** e **N. persone**.
L'unico input di effort è **gg Back office**.

Le colonne giornate esprimono **giornate-uomo aggregate**, non giorni di calendario:
3 consulenti per mezza giornata dal cliente = **1,5** giornate front.

## Da compilare

1. **Nomi** mancanti nel foglio Ruoli (Project Manager e i due junior)
2. Verifica delle **giornate back office**: sono stime, il contratto non le quantifica
3. La **settimana ISO di partenza** nel foglio Timeline, quando il kick-off è schedulato

Tariffe (95 / 85 / 75 €/h), ore di viaggio, pedaggi e vitto sono già compilati.

## Costo delle trasferte

Il blocco nel Riepilogo quantifica un costo **già incluso** nei € 27.000 dell'Offerta
(§ 10.1, 6 incontri in presenza). Si somma al costo delle giornate per ottenere il costo
pieno, da cui si ricava il margine. Semplificazione: i viaggi in auto sono calcolati come
`persone-sessione ÷ persone per auto` arrotondato per eccesso — esatto quando la squadra
riempie l'auto, approssimato per eccesso altrimenti.

## Identità visiva

Segue la skill `impresoft-brand` e la ricetta Excel di `references/office-docs.md`.

- **Palette di tre colori**: giallo `#FDC300`, nero, bianco, più grigi puri derivati dal
  nero (`#F4F4F4` `#3A3A3A` `#6E6E6E` `#8C8C8C` `#9B9B9B` `#DCDCDC`). Nessun altro colore.
- **Font**: Manrope. Se non è installato sulle macchine che apriranno il file, cambiare la
  costante `FONT` in cima a `build_wbs.py` con `"Arial"` e rigenerare.
- **Codifica**: il giallo marca le intestazioni delle colonne da compilare e, nella
  Timeline, i 5 workshop — il nucleo erogato. Le altre fasi scalano in monocromia
  (nero = con il cliente, grigio = interna).
- **Nessun colore semantico**: gli stati usano la scala monocroma (fondo nero, testo
  bianco) e le parole.

### Nota tecnica su openpyxl

I fill vanno dichiarati in ARGB a 8 cifre con `start_color` **e** `end_color` (vedi la
funzione `fill()`). openpyxl omette l'attributo colore quando coincide con il proprio
default, e `000000` normalizzato è quel default: un fill nero dichiarato in modo
abbreviato sparisce dai `dxf` della formattazione condizionale e le barre nere del Gantt
non vengono disegnate.

Attenzione anche alle fasce a piena larghezza: se una fascia unita copre una colonna in
cui poi si scrive un totale, la cella è `MergedCell` e non è scrivibile.

### Struttura delle righe

Titolo in riga 1, intestazioni in riga 2, dati dalla riga 3. Le righe di legenda sono
state rimosse: quelle informazioni vivono nei **commenti delle intestazioni di colonna**,
dove non occupano spazio. Se aggiungi righe di testo sopra la tabella, tutti i
riferimenti degli altri fogli si spostano.

### Nota sul ricalcolo

Il file consegnato **non** passa da LibreOffice: il ricalcolo viene eseguito su una copia
usa e getta, perché la riscrittura di LibreOffice inietta nei fogli di stile record fuori
palette e sostituzioni di font. Le formule sono verificate su quella copia; il file
consegnato ricalcola all'apertura in Excel.
