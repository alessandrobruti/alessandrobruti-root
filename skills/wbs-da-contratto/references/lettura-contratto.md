# Leggere un'offerta tecnico-economica

Un'offerta di servizi è organizzata per convincere, non per pianificare. Le informazioni
che servono a una WBS ci sono quasi tutte, ma sparse su capitoli diversi e con un livello
di precisione molto disomogeneo. Questa è la mappa.

## Dove stanno le cose

| Cosa cerchi | Dove sta di solito |
|---|---|
| Elenco delle fasi | Il capitolo «Soluzione proposta» / «Oggetto della fornitura». Ogni sottoparagrafo è una fase |
| Modalità, durata, risorse | Un blocco «Modalità di esecuzione» dentro ogni sottoparagrafo. È la miniera |
| Deliverable | Un blocco «Deliverable» a chiusura di ogni sottoparagrafo |
| Partecipanti | «Risorse coinvolte» — elencate separatamente per fornitore e per cliente |
| Durata complessiva | Il capitolo «Piano temporale». Confrontala con quella dichiarata nel capitolo delle fasi: spesso divergono |
| Prezzo | Il capitolo «Investimento», in tabella |
| Trasferte incluse | Un sottoparagrafo «Spese di trasferta». Dice **quanti** incontri in presenza copre: è il numero con cui quadrare la WBS |
| Milestone di fatturazione | Il capitolo «Fatturazione e pagamento». Serve a sapere quale fase innesca quale fattura |
| Cosa NON è compreso | Un sottoparagrafo «Esclusioni». Leggilo per ultimo ma leggilo: contraddice spesso i deliverable |
| Vincoli di pianificazione | Sparsi nei «Modalità di esecuzione», in frasi tipo «circa 15 giorni dopo» |

## Cosa estrarre per ogni fase

Per ciascun sottoparagrafo del capitolo delle fasi, la specifica vuole:

- **modalità** — in presenza, remoto, o interna. Determina le giornate di viaggio
- **numero di sessioni** — attenzione: «12 ore complessive (2 workshop di circa 6 ore
  ciascuno)» sono **2** sessioni da 6 ore, non una da 12. È l'errore di lettura più comune
  e falsa sia le giornate front sia il conteggio delle trasferte
- **durata in ore** — quella della singola sessione
- **numero di risorse** — contale nell'elenco «Risorse coinvolte» del fornitore
- **descrizione, attività preparatorie, follow-up, input dal cliente, deliverable** —
  riscritti in forma sintetica e professionale, non copiati
- **partecipanti** — utile tenere anche quelli lato cliente: è la lista di chi convocare

## Le cose che il contratto non dice, e che vanno dichiarate come stime

Non sono lacune del contratto: un contratto non deve dirle. Ma la WBS ne ha bisogno,
quindi vanno stimate **e marcate**, nel campo `note` della fase.

- **le giornate di back office** — il grosso dell'effort e la voce su cui si gioca il
  margine. Nessun contratto le quantifica
- **la durata delle interviste** — i contratti fissano quante, non quanto durano
- **le ore di viaggio** — dedotte dalla distanza fra le sedi indicate
- **km, pedaggi, vitto** — dedotti dalle sedi e dalle abitudini aziendali

## Le trappole che si ripresentano

**Un deliverable obbligatorio la cui lavorazione è esclusa dal perimetro.** Il classico:
il report va consegnato in due lingue, e le esclusioni escludono «traduzione di testi»
come servizio di terze parti. Significa che la seconda lingua è a carico interno, con un
costo che non è previsto altrove. Va messo a piano come fase e segnalato.

**Una durata dichiarata due volte con numeri diversi.** «8-10 settimane» nel capitolo
delle fasi, «10 settimane» nel piano temporale. Prendi il piano temporale come riferimento
e segnala l'incoerenza: va sanata prima della firma.

**Un vincolo di sequenza nascosto in una frase.** «Il meeting di presentazione si svolge
circa 15 giorni dopo il workshop» vincola la coda del progetto e spesso rende impossibile
il piano compresso che l'interlocutore ha in testa. Cercali: stanno nei «Modalità di
esecuzione», mai in un capitolo dedicato.

**Attività che il contratto raggruppa e il piano vuole separate.** Interviste dentro un
workshop, roadmap dentro un workshop di data model. Separarle nella WBS è legittimo e
spesso più utile, ma va scritto nella nota: chi confronta piano e contratto deve capire
perché i conti delle fasi non coincidono.

**Risorse specialistiche promesse al cliente.** Se il contratto nomina un Process Analyst
e un Integration Expert su workshop specifici, e il piano prevede genericamente due
consulenti, al cliente è stata promessa una competenza dedicata. Segnalalo: è un rischio
di delivery, non un dettaglio di nomenclatura.

## Estrarre il testo

Le offerte arrivano in `.docx` o `.pdf`. Il contenuto che serve sta sia nei paragrafi sia
nelle tabelle, e le tabelle spesso contengono il prezzo e l'elenco delle attività:

```python
import docx
d = docx.Document(percorso)
for p in d.paragraphs:
    if p.text.strip():
        print(f"[{p.style.name}] {p.text}")
for i, t in enumerate(d.tables):
    for r in t.rows:
        print(" | ".join(c.text.strip() for c in r.cells))
```

Per i PDF usa la skill `pdf`. Leggi il documento **intero** prima di compilare la
specifica: i vincoli di pianificazione e le esclusioni stanno in fondo, e cambiano il
piano.
