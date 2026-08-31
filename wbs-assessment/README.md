# WBS — Progetto di Assessment Go-to-Market

`WBS_Assessment_GTM.xlsx` è il file di lavoro da editare al desktop.
`build_wbs.py` lo rigenera da zero (`python3 build_wbs.py`) — attenzione: sovrascrive
il file, quindi non usarlo dopo aver iniziato a compilare a mano.

## Fogli

| Foglio | Contenuto |
|---|---|
| **WBS** | Le 12 fasi, una per riga. Colonne giallo/blu = da compilare, grigie = calcolate |
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
