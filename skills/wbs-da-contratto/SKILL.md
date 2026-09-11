---
name: wbs-da-contratto
description: Costruisce una Work Breakdown Structure Excel — fasi, giornate-uomo, costi, margine, Gantt settimanale — da un'offerta tecnico-economica, un contratto di servizi o la presentazione che li accompagna. Ricava per ogni fase modalità, sessioni, durata e risorse, poi SI FERMA E CHIEDE i parametri che i documenti non contengono o che non tornano — lavoro interno, tariffe, distanze, pedaggi, pasti — perché senza quelli il margine non si calcola, e aiuta a bilanciare il numero finale fra giornate, tariffe e trasferte. Usa SEMPRE questa skill quando l'utente allega un'offerta, un contratto, un preventivo o un capitolato e chiede una WBS, un piano di lavoro, una pianificazione, un Gantt, «quantificami le giornate», «quanto ci costa» o «che margine ha». Usala anche per verificare se un progetto già venduto sta in piedi, o per far quadrare un conto che non torna. Vale per consulenza, assessment e delivery a corpo.
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

Da qui due conseguenze operative. La prima: **le giornate front office e di viaggio si
calcolano, le giornate di back office si stimano e si dichiarano come stime.** Nel file la
distinzione è strutturale — le prime sono formule, le seconde sono l'unico input di effort.

La seconda: **prima di generare il file, ci si ferma e si chiede.** Le stime che riguardano
il 50% del costo non le fai tu: le propone e le fa correggere chi eroga il progetto. Un
back office silenziosamente inventato è il modo più rapido per far sembrare solido un
progetto che non lo è, e il passo 2 di questo flusso esiste per impedirlo.

## Flusso di lavoro

### 1. Leggi i documenti e ricava la struttura delle fasi

Apri tutto quello che c'è: l'offerta o il contratto, e la presentazione che l'ha
accompagnata se esiste. I due documenti si completano — il contratto fissa ore, sessioni e
prezzo, la presentazione spesso spiega meglio cosa succede dentro ogni fase e con chi, ed è
la fonte migliore per le descrizioni. Quando divergono vince il contratto, ed è una
divergenza da segnalare. `references/lettura-contratto.md` spiega dove guardare e cosa
cercare, comprese le voci che i contratti di servizi nascondono in capitoli diversi. Per ogni fase
ti servono: modalità, numero di sessioni, durata in ore, numero di risorse, deliverable,
partecipanti per parte.

Compila `spec.json` seguendo `scripts/spec.esempio.json`, che è la specifica completa del
progetto Laica e funge da modello. Per ogni campo che il contratto non copre, scrivi la
stima nel campo e dichiarala nel campo `note` della fase: quelle note finiscono in una
colonna del file e sono ciò che protegge chi lo leggerà dopo di te.

### 2. Fermati e chiedi i parametri che mancano

Questo è il passo che distingue una WBS utile da un esercizio. A questo punto hai le ore
erogate al cliente, che sono nel contratto; **non hai niente di ciò che determina il
margine**, perché il lavoro interno, le distanze e le abitudini di trasferta non stanno in
nessun documento.

Non stimare in silenzio e non generare il file: presenta una **scheda parametri** e chiedi.
Un giro solo, con i valori proposti già dentro, così chi risponde può scrivere «ok tranne
il terzo, metti 8».

La scheda ha tre parti, in questo ordine:

**a) La prima lettura del margine con i default.** Due righe: costo pieno e margine. Serve a
far capire perché le domande contano. Se il margine esce al 2% con le tue stime, la
conversazione cambia natura.

**b) I parametri da confermare**, con accanto **da dove viene il valore proposto** — letto
nel contratto, dedotto dagli indirizzi, stima tua. È l'informazione che permette di capire
a quali righe prestare attenzione.

| Parametro | Proposto | Da dove viene |
|---|---|---|
| gg back office, fase per fase | 3,0 / 1,5 / 2,0 … | stima dai deliverable elencati |
| Persone per sessione | 3 | § 6.2 ne indica 2 — da confermare quale vale |
| Tariffe front / back / viaggio | — | non nei documenti |
| Sede degli incontri | cliente | il § 10.1 lascia scegliere fra le due sedi |
| Km a/r e ore di viaggio | 50 km · 1 h | dedotti dagli indirizzi al § 10.1 |
| Pedaggi · vitto | 0 € · 20 € | stima: tratta ordinaria, un pasto a testa |

**c) Le incoerenze del contratto**, che sono una cosa diversa da un parametro mancante: qui
il documento dice due cose in contrasto e non spetta a te scegliere. Durate dichiarate due
volte con numeri diversi, deliverable obbligatori la cui lavorazione è esclusa dal
perimetro, vincoli di sequenza che rendono impossibile la durata venduta.

`references/parametri-da-chiedere.md` ha la lista completa, i default di partenza e il peso
tipico di ciascuna voce. Leggilo: l'ordine in cui chiedi conta più di quante cose chiedi.

**Una cosa da sapere prima di iniziare a chiedere**, perché orienta tutta la conversazione:
su un progetto di consulenza le giornate di back office pesano il 40-50% del costo pieno,
le ore erogate un altro 40-45%, e **pedaggi, chilometri e pasti insieme l'1-3%**. Chi
affronta il problema dal lato delle spese di trasferta sta guardando la leva sbagliata, e
va detto. Le due leve vere sono il lavoro interno e il numero di persone per sessione.

Per le scelte binarie o con poche opzioni — sede degli incontri, se si viaggia insieme, se
il vitto è pieno — conviene usare uno strumento di domanda strutturata se la sessione ce
l'ha: si risponde con un clic. Per i numeri, la tabella è più rapida.

### 3. Genera il file

```bash
python3 scripts/build_wbs.py spec.json --out "NOME_FILE.xlsx"
```

Lo script produce i cinque fogli (WBS, Timeline, Riepilogo, Parametri, Ruoli) con tutte le
formule, la formattazione condizionale e l'identità visiva. Non riscriverlo a mano: è già
stato validato e contiene le correzioni a un paio di trappole di openpyxl che costano
mezz'ora a riscoprire (sono documentate in testa al file).

### 4. Verifica, che è la parte che distingue un file utile da un file plausibile

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

### 5. Bilancia il numero finale

Se il margine non arriva dove deve, non limare le stime finché il file quadra: è il modo
più rapido per consegnare un progetto che non sta in piedi, e il conto si presenta comunque
tre mesi dopo. Usa invece:

```bash
python3 scripts/bilancia.py spec.json --target 0.25
```

Dice quanto pesa ogni voce sul costo pieno e, per ogni leva, **quale valore servirebbe per
arrivare al margine obiettivo tenendo fermo il resto** — «le giornate di back office
dovrebbero scendere da 18,6 a 8,7», «servirebbero 1,5 persone per sessione invece di 3».
Le leve che «non bastano da sola» sono quelle su cui è inutile discutere.

Porta questi numeri alla conversazione, non una stima già aggiustata. La differenza è che
il primo fa prendere una decisione — tagliamo il back office, andiamo in due, rinegoziamo,
accettiamo il margine — mentre il secondo la rimanda a progetto avviato.

### 6. Riporta gli scostamenti, non solo il file

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
- `references/parametri-da-chiedere.md` — cosa chiedere, con i default di partenza e il peso reale di ogni voce sul costo pieno
- `references/controlli.md` — i controlli di quadratura, cosa significa ogni scostamento, e come si legge il margine
- `scripts/build_wbs.py` — il generatore. Legge `spec.json`, scrive il workbook
- `scripts/spec.esempio.json` — specifica completa di un progetto reale, da usare come modello
- `scripts/verifica.py` — i controlli di quadratura in forma eseguibile
- `scripts/bilancia.py` — peso di ogni voce sul costo pieno e, per ogni leva, il valore che serve a centrare il margine obiettivo
