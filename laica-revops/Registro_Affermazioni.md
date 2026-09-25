# Registro delle affermazioni: Laica S.p.A.

Aggiornato al 25/09/2026. Ogni fatto usato nel report e nella richiesta dati è registrato qui con chi lo dice, dove e con quale affidabilità.

## Scala di confidenza

| Livello | Significato | Uso |
|---|---|---|
| **A** | Documentale: bilancio, RNS Strix, contratto firmato, tariffario pubblico | Si usa come fatto |
| **B** | Dichiarato dall'owner competente, per iscritto o citato testualmente nei verbali | Si usa, citando la fonte |
| **C** | Riportato indirettamente (Campione o riassunto IE) o dichiarato da chi non è owner del tema | Va verificato prima di usarlo |
| **D** | Stima o ipotesi Impresoft | Solo come ipotesi esplicita |
| **E** | Conflitto tra fonti, oppure ignoto | Diventa una domanda al cliente |

## Catena delle fonti

```
Cliente (Dolcetta, Zanuso, Bellero, Milich)
   │  4 call registrate su Bluedot (23/03, 01/04, 15/07, 03/08) + telefonata 20/07
   │    → NON accessibili con l'account Bruti ("role doesn't allow"): chiedere condivisione a Campione
   │  Risposte scritte al questionario (email Zanuso 02/09, docx + ZIP)
   │    → allegati non scaricabili via connettore: salvarli su Drive
   ▼
Documenti Impresoft (Drive): Meeting report · SPICED unificato (07/09) · Presentazione 03/08 · Offerta firmata · WBS V1.1
   ▼
Call interne registrate (Transcript/): 21/07 · 31/07 · KOI 25/09
   → Campione legge lo SPICED e lo commenta a memoria: ogni errore di lettura qui è una sorgente di rumore
```

**Regola pratica:** nel KOI i numeri detti a voce da Campione valgono meno dello SPICED; lo SPICED vale meno della risposta scritta del cliente, che vale meno del dato di sistema. Le conclusioni del report salgono di livello solo quando arriva l'export del sistema.

## Registro

Legenda chi: **RD** Dolcetta · **NZ** Zanuso · **EB** Bellero · **MM** Milich · **CAC** Campione · **AB** Bruti · **IE** documento Impresoft.

| ID | Affermazione | Chi | Fonte | Conf. | Nota critica / verifica |
|---|---|---|---|---|---|
| C-01 | Il 95% dei ricavi viene dal B2B (retail/GDO + OEM), il 5% dall'online | RD, NZ | MR 15/07, 03/08; SPICED | B | **Definizione instabile**: al 03/08 Amazon Vendor sta nel 95% (B2B2C); nel KOI (CAC) sta nel 5%. Nessun dato contabile. |
| C-02 | Target 2 anni della divisione CG: 15,3M GBP (7,6 Y1 + 8,0 Y2), pari al 18,5% del fatturato di divisione | NZ (questionario) | SPICED §S | E | 7,6 + 8,0 = 15,6, non 15,3 (lo nota AB nel KOI). Il 18,5% implica una divisione da ~£83m, contro £34m reali (Strix RNS). Va chiarito se è un target incrementale, D2C o altro. |
| C-03 | Quota OEM sul B2B | — | KOI (VM, AB: "non ancora") | E | Ignota. È il dato più importante per pesare i due Revenue Model. |
| C-04 | Gli OEM ricevono prodotti su misura, esclusivi (es. macchine per latte per bambini in UK) | RD | MR 23/03 | B | Coerente con Strix RNS: crescita CG da "leading global baby brand" (A). |
| C-05 | Piano triennale 01/04/2026–31/03/2029: consolidamento di IT, UK, ES ed Est Europa; attivazione di FR e DE; America esclusa | RD | MR 23/03 | B | Nel KOI è "da confermare se biennale o triennale": è triennale. |
| C-06 | L'online dovrà coprire tra il 50% e il 75% del lavoro per centrare gli obiettivi di crescita | RD | MR 23/03 | B | Nessun target numerico sul D2C: va chiesto. |
| C-07 | Budget del piano triennale ed eventuale approvazione della PLC | — | 21/07 (CAC: "da chiedere") | E | Nuovo CEO Strix Rainforth dal 13/07/2026, revisione strategica in corso (A web): rischio sul budget. |
| C-08 | Molti OEM non sanno che Laica produce white label; promozione OEM pari a zero | RD | MR 03/08 (citazione testuale) | B | — |
| C-09 | Il brand è noto in Italia per le caraffe, poco per gli altri prodotti e poco all'estero | RD via CAC | 21/07, KOI | C | Misurabile con share of search e brand search (K41). |
| C-10 | Il fatturato Amazon è quasi tutto Amazon.it; gli altri paesi valgono poco o niente | CAC ("notizia certa") | KOI | C | Account aperti in FR, DE, ES, PL "per vedere se si vende" (NZ/RD, MR 01/04, B). |
| C-11 | Amazon Seller costa il 18% di fee più il 20-22% di FBA, stoccaggio e resi, cioè 38-40% | NZ | MR 15/07 (testuale "oltre il 20%"); SPICED | B | **Nel KOI è diventato "20-22% in totale" (errore di lettura).** Lo SPICED dice ">40%". Tariffa pubblica Amazon.it Casa e cucina: 15%, 8% sotto €20 (A): il 18% va riconciliato. Esclude gli ADS. |
| C-12 | Tasso di resi Amazon Seller 4,13% | NZ (questionario) | SPICED §P | B | Metodo di calcolo ignoto (unità o valore). |
| C-13 | Budget AMS 595.000 €/anno, circa 50.000 €/mese tra Vendor e Seller | NZ (questionario) | SPICED §P | B | Nel KOI CAC era incerto. È circa il 48% di un online da ~€1,25m: **incoerente con il mix 95/5**, a meno che il Vendor sia grande. Verificare. |
| C-14 | Il team Sales dedica 40 h/mese alle promozioni Amazon; la gestione dei canali assorbe il 40% del tempo di Bellero | NZ/EB | SPICED; MR 15/07 | B | Due misure diverse: 40 h/mese ≈ 0,25 FTE sales; il 40% di EB = 0,4 FTE marketing. La presentazione usa 0,25 FTE mkt × RAL 40k = €13,5k. |
| C-15 | HubSpot usato come repository: circa 20.000 anagrafiche, di cui **15.820 con consenso esplicito**; moduli Sales e Marketing non attivi; solo mercato IT | EB (01/04), NZ (questionario) | MR 01/04; SPICED | B | **Il KOI dice "15.000, la maggior parte senza consenso": falso.** Il 79% ha il consenso. Licenza HubSpot ignota. |
| C-16 | Shop ricambi: abbandono cumulato 67,1% (47% carrello→checkout, 37,9% checkout→acquisto); 16 ordini/mese | NZ/EB (questionario) | SPICED | B | Conti coerenti. "Conversione 0,08% sul database" = 16/20.000: metrica senza senso, non usarla. |
| C-17 | Customer care 2025: 6.435 ticket (4.129 email, 2.228 form, 78 Facebook); prima risposta 3,3 gg via email e 9,1 gg via form; 80% FAQ ricorsive (39,5% sottovuoto assistenza, 26,4% ricambi sottovuoto, 21% caraffe); up-sell/cross-sell 1-2% | NZ (questionario); "1 risorsa quasi totalmente impiegata" (NZ, testuale) | SPICED; MR 15/07 | B | KOI: 6.465 e "9,9" (errori). 39,5 + 26,4 + 21 = 86,9 ≠ 80. Strumento di ticketing ignoto. |
| C-18 | Riattivando il 25% del database a 30 € si ottengono 150 k€/anno | IE (CAC) | 31/07; presentazione | D | 25% × 20.000 × 30 = 150k, ma la base con consenso è 15.820 (→ 118,6k). Il tasso del 25% non ha base. AB (31/07): "ipotesi meno ancorata". |
| C-19 | Deal OEM tra 250 k€ e 4 M€, chiusura in 1-3 anni | NZ/MM (questionario) | SPICED §P | B | Pipeline e win rate ignoti. |
| C-20 | Criterio del CFO: payback inferiore a 24 mesi | NZ (questionario) | SPICED §D | B | — |
| C-21 | Leva Amazon→D2C: omaggio o estensione di garanzia in confezione, con raccolta email | AB (esempio "punte trapano") | MR 01/04; 21/07 | D | NZ l'ha collegata a sottovuoto e sacchetti (B). Serve il dato di registrazione prodotto (K38). |
| C-22 | Philips (licenziatari) ha comunicato un filtro anti-PFAS 6 mesi prima di Laica, con tecnologia inferiore | RD | MR 23/03 | C | "Inferiore" è un giudizio di parte. Il ritardo di 6 mesi è confermato più volte, ma senza date. |
| C-23 | Integrazione ERP↔Amazon Seller | NZ: "l'ERP comunica live... quasi un anno di implementazione" (03/08) | MR 03/08 | E | Lo SPICED e il KOI dicono invece "script CSV giornalieri". Vendor: CSV scaricati da Vendor Central (EDI escluso per costi). Chiedere a Sella. |
| C-24 | Progetto PIM avviato circa 5 anni fa e interrotto per mancanza di un owner | NZ/EB | MR 15/07 | B | **Nel KOI è confuso con il fallimento Salesforce di Milich in Pastiglie Leone** (21/07, CAC): sono due storie diverse. |
| C-25 | Un progetto di AI predittiva sulle vendite si è arenato per dati sporchi ("l'algoritmo ha le allucinazioni") | cliente (probabilmente NZ) | MR 15/07 (testuale) | B | Prova diretta del problema di data quality. |
| C-26 | Lo shop è WordPress + WooCommerce, manutenzione 2.000 €/anno per 5 h/mese | NZ (questionario) | SPICED §S | B | Nel KOI CAC oscilla tra Shopify e WordPress: risolto. |
| C-27 | Vendita D2C di hardware | SPICED: "non prevista fino al FY29/30"; RD 20/07: "imminente estensione dell'e-commerce a tutti i prodotti"; MM 03/08: "punto molto urgente" | SPICED; MR 20/07, 03/08 | E | Conflitto aperto. Domanda chiave da fare a RD e MM insieme. |
| C-28 | Il D2C vale "poche centinaia di migliaia di euro" | RD/NZ | MR 03/08 | E | Con 16 ordini/mese a ~30 € si arriva a ~€6k/anno. Forse include il Seller. Chiarire (K43). |
| C-29 | Laica fa white label anche per Philips | CAC ("che ne so") | 21/07 | D | Esempio inventato. Non usarlo. |
| C-30 | Rete commerciale Italia: 5 agenzie per l'elettronica di consumo, 5 per il DIY, rapporto diretto con Unieuro e MediaWorld. UK: key account diretti (Tesco). ES: Sales Manager "Pedro" con agenzie. Altri paesi: distributori. Nessun magazzino 3PL | MM | MR 03/08 | B | — |
| C-31 | Canali ordini: email ad agenti o alla sede, portale B2B proprietario per i distributori, Seller Central con approvazione manuale | MM/NZ | MR 03/08 | B | Il portale B2B non è citato nel KOI: aggiungerlo al tech stack. |
| C-32 | Est Europa presidiata anche su Allegro ed eMAG | cliente | MR 15/07 | B | Il KOI dice "solo Amazon": incompleto. |
| C-33 | Agenzia MoMA per le campagne, non soddisfacente | RD | MR 23/03 | B | Budget e canali dell'agenzia ignoti. |
| C-34 | Esiste un'app mobile, sviluppata in Cina, per i device connessi | CAC | 21/07 | C | Dati app (registrazioni, utenti) potenzialmente preziosi per il Bow Tie Growth. |
| C-35 | Nessun PIM, DAM, CDP né chat; i dati di prodotto stanno in database paralleli; non conformità rilevate da un auditor ISO | RD, EB | MR 23/03, 15/07 | B | — |
| C-36 | Roadmap 2026-27: robot pulisci vetri, cooking blender, sottovuoto VT3206/VT3260, bilance smart, dental/medical, filtri slim Mg/limescale, PFAS | RD | SPICED; MR | B | — |
| C-37 | "Cinque persone prendono decisioni non coordinate su Amazon" | cliente | MR 01/04 | B | Mappare i ruoli nel WS#3. |
| C-38 | Il 99% del budget di comunicazione va sul brand Laica | RD | MR 23/03 | B | — |
| C-39 | Rischio persone: le figure operative "Volga" e "Richard" portano il tavolo sul tattico | NZ | MR 01/04 | B | Nomi non mappati nell'organigramma del KOI. |
| C-40 | Atteggiamenti: MM guardingo all'inizio, EB collaborativa (secondo CAC); EB "difensore dello status quo" su HubSpot e Amazon (MR 01/04) | CAC / IE | 21/07; MR 01/04 | D | **Due letture opposte su EB.** Sono opinioni, non fatti. |
| C-41 | Il responsabile IT di gruppo sta uscendo | RD | MR 23/03 | C | Non è chiaro se sia Strix o Laica. Sella resta referente? |
| C-42 | Il cliente ha chiesto di non usare la parola "RevOps" | RD | MR 20/07 | B | Vale anche per i deliverable destinati agli operativi. |
| C-43 | Billi venduta per £110m (EV); Strix passa a cassa netta di £38,7m al 31/03/2026 | — | RNS Strix (web) | A | RD dice "venduta al triplo" (C): 3x non verificato. |
| C-44 | Nuovo CEO Strix Andy Rainforth dal 13/07/2026, revisione strategica in corso | — | RNS (web) | A | Coerente con la "finestra di 6 mesi" citata da RD a marzo. |
| C-45 | Strix Consumer Goods: £35,4m (FY22), £32,3m (FY23), £31,4m (FY24), £34,4m (12 mesi a 03/2026) | — | RNS (web) | A | Crescita recente trainata dall'OEM, non dai brand. |
| C-46 | Laica S.p.A.: ricavi €20,4m e EBITDA rettificato €2,9m (2019) | — | RNS di acquisizione (web) | A | — |
| C-47 | Laica S.p.A.: ricavi €24,34m (2023) e €25,15m (2024); utile €1,59m (2024) | — | Aggregatori di bilancio (web) | C | Serve il bilancio depositato (K01). |
| C-48 | Acquisizione Strix 10/2020: €19,6m più earn-out fino a €12m | — | RNS (web) | A | — |
| C-49 | Laica dichiara oltre 100 brevetti; il bi-flux universal con "LAICA-Key" è compatibile con Brita | — | laica.it (web) | A | In Germania Brita ha vinto contro Aqua Select sulla pubblicità della compatibilità (B): rischio legale. |
| C-50 | Referral fee Amazon.it Casa e cucina 15% (8% sotto €20) | — | Tariffario Amazon (web) | A | Base per riconciliare C-11. |
| C-51 | Contratto firmato il 28/08/2026: €27.000 una tantum, trasferte per 6 incontri incluse, 50% alla firma e 50% alla consegna del PDF | — | Offerta firmata CX260803 | A | — |
| C-52 | WBS V1.1: 37,41 gg; costo pieno €26.996,5 contro ricavo €27.000, **margine 0,0%** | — | WBS V1.1 | A | Nessun buffer: la raccolta dati deve essere efficiente (una sola richiesta, completa). |
| C-53 | Durata "8-10 settimane" (§6) contro "10 settimane" (§8) | — | Offerta v1.1 | A | Incoerenza contrattuale segnalata nella WBS. |
| C-54 | Kick-off cliente il 02/10 (invito inviato); dati entro il 09/10; analisi il 15/10; review il 16/10; workshop dalla settimana del 19/10 | AB | KOI; Gmail | A | La WBS V1.1 colloca l'1.1 in W41-42 (5-18/10): allineare. |
| C-55 | Sella (IT) "senza grandi idee" | CAC (mai incontrato) | KOI | D | Opinione di seconda mano, da non usare. |
| C-56 | Target di crescita D2C, CAC, LTV, tasso di riacquisto dei filtri | — | Questionario §2 | E | Non misurati (SPICED: "non storicizzati"). |
| C-57 | Importo AMS e fee rispetto al fatturato Amazon totale | — | — | E | Manca il GMV Amazon: senza non si calcolano TACoS né incidenza. |
| C-58 | Fatturato OEM, retail e Vendor in euro | — | — | E | Primo export da chiedere (K01). |

## Errori da non propagare (KOI → correggere nel kick-off)

| Detto nel KOI | Fonte corretta |
|---|---|
| "Amazon costa il 20-22%" | 18% di fee + 20-22% di FBA = 38-40% (NZ), più gli ADS |
| "15.000 contatti, la maggior parte senza consenso" | 20.000 anagrafiche, di cui 15.820 con consenso |
| "6.465 ticket", "9,9 giorni" | 6.435 ticket; 9,1 giorni per i form |
| "E-commerce su Shopify" | WordPress + WooCommerce |
| "PIM fallito perché calato dall'alto" | PIM interrotto per mancanza di un owner (NZ); "calato dall'alto" è il caso Salesforce di Milich in Pastiglie Leone |
| "Il 5% = Amazon Seller e Vendor" | Al 03/08 il Vendor sta nel 95% |
| "Piano biennale o triennale?" | Triennale, 04/2026–03/2029 |
