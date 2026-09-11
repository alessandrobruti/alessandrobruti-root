# I controlli di quadratura, e come si leggono

`scripts/verifica.py` esegue questi controlli. Nessuno di essi, se fallisce, indica
necessariamente un errore del file: metà delle volte indica una crepa nel contratto. Qui
sta la differenza.

## Quadratura del costo delle giornate

Il costo totale deve coincidere con la somma dei costi per tipologia di ora. Se non
coincide, l'errore è **tuo**: una formula sbagliata o un intervallo che non copre tutte le
righe. È il solo controllo di questa lista che non ammette interpretazioni.

## Sessioni in presenza vs trasferte incluse

Il contratto dichiara quanti incontri in presenza copre con le spese di trasferta. La WBS
lo ricalcola dalle fasi. Se i due numeri divergono:

- **la WBS ne ha di più** — hai messo in presenza una fase che il contratto prevede in
  remoto, oppure hai contato male le sessioni di una fase multipla. Ricontrolla il
  contratto prima di toccare il file
- **la WBS ne ha di meno** — o hai perso una fase, o il contratto conta un incontro che nel
  piano non esiste più
- **coincidono** — è la conferma più forte che hai letto bene il capitolo delle fasi

## Lead time vs durata dichiarata

Lo scostamento dice qualcosa di diverso a seconda del segno.

- **effettivo maggiore del dichiarato** — il piano non sta nella durata venduta. Cause
  tipiche: una coda di lavorazione che il contratto colloca dopo l'ultimo incontro, o un
  vincolo di sequenza che allunga la fine. Non si risolve comprimendo a caso: si decide se
  la coda è dentro o fuori la durata dichiarata, e si dice
- **effettivo minore del dichiarato** — il piano è più corto di quanto venduto. Non è un
  errore, ma verifica che non sia dovuto a fasi messe in parallelo che in realtà non
  possono sovrapporsi, o a un vincolo di sequenza che ti è sfuggito

## Margine e tariffa media implicita

Sono le due letture che contano davvero, e vanno lette insieme.

**La tariffa media implicita** è l'investimento diviso il totale delle ore. Se esce sotto
la più bassa delle tariffe impostate, il progetto non regge nemmeno se lo eroghi con le
figure più economiche. Se esce vicina a quella più bassa, regge solo a condizione che il
mix di seniority sia molto più leggero di quello pianificato.

**Il margine** è l'investimento meno il costo pieno, trasferte incluse. Un margine vicino
a zero non è un errore di calcolo: è il progetto che è stato venduto così. Le leve, in
ordine di efficacia:

1. **le giornate di back office** — sono stime, sono la voce più grossa, e sono quelle su
   cui chi eroga ha davvero un'opinione. Chiedile invece di difendere le tue
2. **il numero di persone per sessione** — una persona in più su un workshop da due
   sessioni da sei ore costa una giornata e mezza front. È la leva più rapida
3. **le fasi fuori perimetro contrattuale** — una sessione tenuta a piano per scelta è una
   giornata regalata: legittima, se consapevole
4. **la sede degli incontri** — se il contratto lascia scegliere fra la sede del cliente e
   la propria, gli incontri in casa azzerano la trasferta

Quello che **non** è una leva: ridurre le ore erogate al cliente. Quelle sono contrattuali.

## Fasi con incontri incomplete

Se una fase ha un numero di incontri ma non la durata o le persone, le giornate front
escono a zero **senza segnalare nulla**. È il tipo di buco che rende un file plausibile e
sbagliato. Il controllo esiste per questo.

## Back office valorizzato

Non è un controllo di correttezza ma un promemoria: se il totale di back office è zero, il
file dice che il progetto si eroga senza lavoro interno. Non capita mai.

## Cosa fare dopo i controlli

Gli scostamenti che restano vanno **riportati a chi ti ha chiesto il lavoro**, non
corretti in silenzio. Un file che quadra perché hai limato le stime finché quadrava è
peggio di un file che non quadra e lo dice: il secondo fa prendere una decisione, il primo
la rimanda al momento in cui il progetto è già partito.
