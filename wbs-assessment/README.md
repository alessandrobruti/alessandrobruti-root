# WBS — RevOps Assessment B2B e B2C · Laica S.p.A.

`WBS_Assessment_GTM.xlsx` è il file di lavoro da editare al desktop.
`build_wbs.py` lo rigenera da zero (`python3 build_wbs.py`) — sovrascrive il file,
quindi non usarlo dopo aver iniziato a compilare a mano.

**Fonte:** Offerta `CX260803_Off_Laica_RevOps assessment B2B e B2C` v1.1 del 28 ago 2026.
Modalità, ore e numero di sessioni di ogni fase sono letti dal capitolo 6 del contratto.

## Fogli

| Foglio | Contenuto |
|---|---|
| **WBS** | Le 11 fasi, una per riga. Intestazione gialla = da compilare, celle grigie = calcolate |
| **Timeline** | Gantt per settimana, si disegna da sé dalle settimane impostate nella WBS |
| **Riepilogo** | Impegno, costo giornate, costo trasferte, costo pieno e margine, lead time, punti aperti |
| **Parametri** | Ore/giornata, tre tariffe orarie, ore di viaggio, investimento, costi di trasferta |
| **Ruoli** | Mappatura ruolo → nome del team di progetto |

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

1. Le tre **tariffe orarie** (front / back / viaggio) nel foglio Parametri — finché sono vuote tutti i costi valgono 0
2. **Vitto** €/persona/trasferta, e i pedaggi se il percorso prevede autostrada
3. **Durata** dell'incontro di review con il cliente (fase 8), oggi vuota
4. **Nomi** del team nel foglio Ruoli
5. Verifica delle **giornate back office**: sono stime, il contratto non le quantifica

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

### Nota sul ricalcolo

Il file consegnato **non** passa da LibreOffice: il ricalcolo viene eseguito su una copia
usa e getta, perché la riscrittura di LibreOffice inietta nei fogli di stile record fuori
palette e sostituzioni di font. Le formule sono verificate su quella copia; il file
consegnato ricalcola all'apertura in Excel.
