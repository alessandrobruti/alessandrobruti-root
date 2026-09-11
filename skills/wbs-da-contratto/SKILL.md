---
name: wbs-da-contratto
description: Costruisce una Work Breakdown Structure Excel — fasi, impegno in giornate-uomo, costi, margine, Gantt settimanale — partendo da un'offerta tecnico-economica o da un contratto di servizi. Legge il contratto, ricava da ogni fase modalità, numero di sessioni, durata e risorse, calcola le giornate front office e di viaggio con formule vive, e produce un file a cinque fogli editabile al desktop. Usa SEMPRE questa skill quando l'utente allega un'offerta, un contratto, un preventivo o un capitolato e chiede una WBS, un piano di lavoro, una pianificazione, un Gantt, "quantificami le giornate", "quanto ci mettiamo", "quanto ci costa questo progetto" o "che margine ha". Usala anche quando chiede di verificare se un progetto già venduto sta in piedi — se le giornate, il lead time e il prezzo sono coerenti fra loro — o di confrontare un piano esistente con quello che il contratto dice davvero. Vale per progetti di consulenza, assessment, system integration e delivery a corpo, non solo RevOps.
---

# WBS da contratto

Un'offerta firmata è il documento più affidabile che esista su un progetto, e allo stesso
tempo dice solo metà delle cose che servono a pianificarlo. Questa skill serve a estrarre
la metà che c'è, marcare con chiarezza la metà che manca, e restituire un file su cui il
project manager può lavorare senza rifare i conti a mano.

## Il principio che regge tutto

**Un contratto quantifica il tempo erogato al cliente. Non quantifica quasi mai il lavoro
interno.** Ti dirà che il workshop dura 6 ore, che le sessioni in presenza sono sei e che
l'investimento è di 27.000 euro. Non ti dirà quante giornate servono per scrivere il
report finale.

Da qui la regola operativa: **le giornate front office e di viaggio si calcolano, le
giornate di back office si stimano e si dichiarano come stime.** Nel file questa
distinzione è strutturale — le prime sono formule, le seconde sono l'unico input di
effort — e va detta anche a voce quando consegni. Un back office silenziosamente inventato
è il modo più rapido per far sembrare solido un progetto che non lo è.

## Flusso di lavoro

### 1. Leggi il contratto e ricava la struttura delle fasi

Apri il documento (`references/lettura-contratto.md` spiega dove guardare e cosa cercare,
comprese le voci che i contratti di servizi nascondono in capitoli diversi). Per ogni fase
ti servono: modalità, numero di sessioni, durata in ore, numero di risorse, deliverable,
partecipanti per parte.

Compila `spec.json` seguendo `scripts/spec.esempio.json`, che è la specifica completa del
progetto Laica e funge da modello. Per ogni campo che il contratto non copre, scrivi la
stima nel campo e dichiarala nel campo `note` della fase: quelle note finiscono in una
colonna del file e sono ciò che protegge chi lo leggerà dopo di te.

### 2. Genera il file

```bash
python3 scripts/build_wbs.py spec.json --out "NOME_FILE.xlsx"
```

Lo script produce i cinque fogli (WBS, Timeline, Riepilogo, Parametri, Ruoli) con tutte le
formule, la formattazione condizionale e l'identità visiva. Non riscriverlo a mano: è già
stato validato e contiene le correzioni a un paio di trappole di openpyxl che costano
mezz'ora a riscoprire (sono documentate in testa al file).

### 3. Verifica, che è la parte che distingue un file utile da un file plausibile

Una ricalcolata pulita dimostra che le formule *valutano*, non che i numeri sono *giusti*.
Fai girare i controlli di quadratura:

```bash
python3 scripts/verifica.py "NOME_FILE.xlsx" --spec spec.json
```

Confronta le sessioni in presenza calcolate con il numero di trasferte che il contratto
dichiara incluse, il lead time con la durata dichiarata, e la tariffa media implicita con
le tariffe impostate. `references/controlli.md` spiega cosa significa ogni scostamento —
quali sono errori tuoi e quali sono crepe nel contratto.

Il ricalcolo va fatto **su una copia**: LibreOffice, riscrivendo il file, inietta nei fogli
di stile record fuori palette e sostituzioni di font. Si consegna l'originale.

### 4. Riporta gli scostamenti, non solo il file

Il valore di questo lavoro non è il file: è la lista di cose che non tornano. Consegna
sempre, insieme al file, un riepilogo in chat di:

- **cosa il contratto dice e l'interlocutore crede diverso** — capita quasi sempre, ed è
  la parte che gli fa cambiare il piano
- **le incoerenze interne al contratto** — durate dichiarate due volte con numeri diversi,
  deliverable obbligatori la cui lavorazione è esclusa dal perimetro, vincoli di
  pianificazione impossibili da rispettare
- **il margine e la tariffa media implicita** — se il prezzo venduto non regge le giornate
  necessarie, è questo il momento in cui si scopre, non a progetto avviato

Ordina per impatto economico, non per ordine di lettura del contratto.

## Le convenzioni del file, e perché sono quelle

**Le colonne giornate esprimono giornate-uomo aggregate, non giorni di calendario.** Tre
consulenti per mezza giornata dal cliente fanno 1,5 giornate front, non 0,5. La colonna
serve a sommare sforzo e costo; la composizione della squadra si descrive a parole nella
colonna Delivery. Senza questa convenzione la stessa colonna finisce per contenere due
grandezze diverse e nessun totale è più leggibile.

**Front office è il tempo erogato al cliente, incontri in remoto inclusi.** La
distinzione utile è cliente-vs-interno, non presenza-vs-remoto: un workshop di 12 ore in
videocall non è back office. La presenza fisica conta altrove, cioè nelle giornate di
viaggio e nel costo delle trasferte.

**Le settimane della WBS sono relative: il progetto parte sempre dalla settimana 1.** Il
lead time si misura da lì e non cambia se il kick-off slitta. La settimana di calendario
si imposta in una cella sola, in testa alla Timeline: le etichette e la fascia dei mesi si
spostano, le barre no. È così che si fanno le proiezioni — «se partiamo a inizio ottobre
chiudiamo a fine dicembre» — senza rimettere mano a ventidue celle.

**Il giallo va alla serie che porta il messaggio.** Nella Timeline è il blocco di workshop,
cioè il nucleo erogato; le altre fasi scalano in monocromia. Questo perché la palette
Impresoft ha un solo colore d'accento, e sprecarlo su una categoria che non è il punto del
foglio è il modo più rapido per rendere il Gantt illeggibile.

## Identità visiva

Il file segue la skill `impresoft-brand`. Lo script bundled la implementa già: palette di
tre colori più grigi puri derivati dal nero, font Manrope con fallback Arial, nessun colore
semantico. **Carica comunque `impresoft-brand` prima di modificare lo script**, perché le
regole di contrasto del giallo sono controintuitive e la ricetta Excel
(`references/office-docs.md` di quella skill) ha vincoli precisi su intestazioni, righe
alternate e formattazione condizionale.

Un punto che si sbaglia facilmente: **il brand non ha verde «ok» e rosso «errore»**. Gli
stati si codificano con la scala monocroma — fondo nero, testo bianco — e con etichette in
parole. Nel file lo scostamento di lead time funziona così.

## Quando il contratto e l'interlocutore non concordano

Succede regolarmente, e non è un problema da risolvere in silenzio. Il contratto è ciò che
è stato venduto; l'interlocutore sa cose che nel contratto non ci sono e a volte ricorda
male ciò che c'è. Quando divergono:

- **metti a piano quello che dice l'interlocutore**, se è una sua decisione consapevole —
  è lui che eroga il progetto
- **segnala lo scostamento nella colonna Note della fase**, con il riferimento al paragrafo
  del contratto, così chi apre il file mesi dopo capisce perché il piano dice una cosa e
  l'offerta un'altra
- **quantifica il costo dello scostamento** quando ne ha uno: «tre persone dove il
  contratto ne prevede due, sul workshop da due sessioni fa una giornata e mezza front in
  più» è un'informazione su cui si decide, «il contratto dice diversamente» non lo è

Una fase fuori perimetro contrattuale tenuta a piano su richiesta va marcata come tale a
lettere chiare. Non è pedanteria: è la differenza tra una giornata che qualcuno ha deciso
di regalare e una che si è persa per distrazione.

## File di riferimento

- `references/lettura-contratto.md` — dove stanno le informazioni in un'offerta tecnico-economica, cosa estrarre e le voci che si trovano solo cercandole
- `references/controlli.md` — i controlli di quadratura, cosa significa ogni scostamento, e come si legge il margine
- `scripts/build_wbs.py` — il generatore. Legge `spec.json`, scrive il workbook
- `scripts/spec.esempio.json` — specifica completa di un progetto reale, da usare come modello
- `scripts/verifica.py` — i controlli di quadratura in forma eseguibile
