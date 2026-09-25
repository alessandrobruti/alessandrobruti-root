# -*- coding: utf-8 -*-
"""Albero KPI Laica (modello Bow Tie) + richiesta dati.
Fonte unica per KPI_Tree.md e Data_Request_LAica.xlsx.

Motion: RET = B2B retail/distribuzione (sell-in) · OEM = white label · AMZ = Amazon Seller/Vendor · D2C = e-commerce proprio · ALL = trasversale
Stadi Bow Tie: 1 Awareness · 2 Education · 3 Selection · 4 Commit · 5 Onboarding · 6 Impact · 7 Growth · X = trasversale (economics / dato)
Livello: L1 Business · L2 Servizio · L3 Operativo
Stato: DIR = sistema noto → richiesta diretta · DISC = sistema/dato ignoto → domanda di discovery
"""

# Owner (da KOI 25/09, fonte C: riportato da Campione)
OWN = {
    "CFO": "Nicolò Zanuso (CFO)",
    "AMM": "Alberto Zagnoli / Sara Cacace (Amministrazione)",
    "COMM": "Marco Milik (Resp. Commerciale)",
    "MKT": "Elisa Bellero (Marketing Mgr)",
    "AMZ": "Clara Contin (Commerciale Amazon)",
    "CS": "Stefania Zen (Customer Service)",
    "CSA": "F. Ghirardello / E. Beggiato (CS Amazon)",
    "IT": "Danilo Sella (IT)",
    "CEO": "Riccardo Dolcetta (CEO Consumer Goods)",
}

# (id, motion, stadio, livello, kpi, formula, perche, dato, granularita, sistema, owner, stato, richiesta, prio)
KPI = [
 # ---------------- L1 BUSINESS / ECONOMICS ----------------
 ("K01","ALL","X","L1","Ricavi netti per canale e motion","Σ fatturato netto (dopo sconti, premi, NC) per motion RET/OEM/AMZ-V/AMZ-S/D2C/Export",
  "Mix 95/5 B2B/online mai documentato; % OEM ignota (C-01, C-03)","Righe fattura: data, cliente, canale, paese, SKU, qty, prezzo lordo/netto, sconti, NC","Riga fattura · mensile · 2023→YTD 2026","ERP JGalileo","AMM","DIR",
  "Export JGalileo righe fatture+NC 01/2023→oggi (CSV): n.doc, data, cod.cliente, rag.soc., paese, canale/agente, cod.SKU, desc., famiglia, qty, prezzo lordo, sconti, netto, costo standard","M"),
 ("K02","ALL","X","L1","Margine lordo (contributo) per canale / famiglia / cliente","(Ricavo netto − COGS − costi variabili canale) / ricavo netto",
  "Erosione margine Amazon 20-22% vs 38-40% (dato conflittuale, C-11)","COGS landed per SKU (FOB Cina + dazi + nolo), premi/contributi trade, fee canale","SKU · anno · 2024-2026","ERP JGalileo + controllo di gestione (Excel?)","CFO","DISC",
  "Esiste un conto economico per canale/cliente? Dove si calcola il costo landed per SKU (ERP o Excel)? Chi lo aggiorna? Ce lo condividete per 2024-2025?","M"),
 ("K03","ALL","X","L1","Crescita ricavi YoY per motion e paese","(Ricavi t / Ricavi t-1) − 1","Obiettivo piano triennale: consolidamento EU + espansione FR/ES (C-05)","Da K01","Mensile · 3 anni","ERP JGalileo","AMM","DIR","Incluso in export K01 (campo paese destinazione + paese cliente)","M"),
 ("K04","ALL","X","L1","Target piano triennale 2026-2028 (ricavi, mix, paesi, budget)","Target vs actual","Piano da approvare dalla PLC; nessun target numerico noto (C-06, C-07)","Piano industriale / budget 2026-2028","Anno","Documento CFO/CEO","CEO","DISC",
  "Potete condividere il piano triennale (anche bozza): ricavi target per canale e paese al 2028, peso D2C atteso, budget marketing/tech stanziato, stato approvazione Strix PLC?","M"),
 ("K05","ALL","X","L1","Concentrazione clienti (top-10 share, Pareto)","Ricavi top-10 clienti / ricavi totali","Dipendenza da GDO/retail e pochi OEM (ordini 250k-4M€)","Da K01","Cliente · anno","ERP JGalileo","AMM","DIR","Incluso in export K01","M"),
 ("K06","ALL","X","L1","Costo del servizio clienti per ordine / per ricavo","Costo CS (FTE×costo) / n. ordini o ricavi","CS = 1 FTE assorbito da ricorsive (CFO, C-17)","FTE e costo CS, n. ordini per canale","Anno","HR/Amministrazione","CFO","DISC","Quante persone (FTE) lavorano su CS diretto e CS Amazon, e con quale costo aziendale medio?","S"),
 ("K07","ALL","X","L1","Payback / ROI iniziative RevOps (≤24 mesi)","Beneficio cumulato / investimento","Richiesta esplicita CFO ROI ≤24 mesi (C-20)","Baseline K01-K60 + costo iniziative","Trimestre","Calcolato","CFO","DISC","Qual è il tasso/soglia con cui valutate un investimento (payback, IRR)? Che costi includete (licenze, persone, agenzie)?","S"),

 # ---------------- RETAIL / DISTRIBUZIONE (B2B sell-in) ----------------
 ("K10","RET","3","L3","Nuovi punti vendita / nuovi clienti B2B attivati","N. clienti con primo ordine nel periodo","Strategia espansione FR/ES richiede nuovi distributori","Anagrafica clienti + data primo ordine","Mensile · 3 anni","ERP JGalileo","COMM","DIR","Export anagrafica clienti JGalileo: cod., rag.soc., paese, tipologia (GDO/elettronica/farmacia/grossista/OEM), agente, data creazione, stato","M"),
 ("K11","RET","4","L1","Sell-in per cliente / insegna / famiglia","Σ ricavi netti per cliente × famiglia","Rete fisica = canale principale (KOI)","Da K01","Mensile","ERP JGalileo","COMM","DIR","Incluso in K01","M"),
 ("K12","RET","6","L2","Sell-out e rotazione a scaffale (ove disponibile)","Pezzi venduti al consumatore / punti vendita / settimane","Senza sell-out non si chiude sell-in↔consumatore (Bruti 21/07)","Report sell-out da insegne (EDI/portali) o panel (GfK/NIQ)","Settimanale/mensile","Portali retailer / GfK?","COMM","DISC","Ricevete dati di sell-out da Unieuro/MediaWorld/GDO o panel GfK/NielsenIQ? Con che frequenza e formato? Ce ne date 12-24 mesi?","S"),
 ("K13","RET","6","L2","OTIF / fill rate verso retailer","Righe consegnate complete e puntuali / righe ordinate","Produzione in Cina → rischio stock-out e lead time","Ordini cliente (data richiesta vs data consegna, qty ord vs sped)","Riga ordine · 2024-2026","ERP JGalileo","AMM","DIR","Export ordini clienti JGalileo con date richiesta/conferma/spedizione e qty ordinata/evasa","S"),
 ("K14","RET","6","L2","Resi e NC da trade (% ricavi)","Valore NC reso / ricavi","Resi 4,13% (riferito ad Amazon? da chiarire, C-12)","Note credito con causale","Mensile","ERP JGalileo","AMM","DIR","Incluso in K01 (causale NC obbligatoria)","S"),
 ("K15","RET","7","L1","Retention clienti B2B e NRR","Ricavi anno t da clienti attivi t-1 / ricavi t-1 degli stessi","Rischio perdita spazio a scaffale vs Brita/Philips","Da K01","Annuale","ERP JGalileo","COMM","DIR","Incluso in K01","M"),
 ("K16","RET","X","L1","Trade spend (premi, contributi, promo) % sell-in e ROI promo","Trade spend / sell-in; Δvolume promo / costo promo","Promozioni gestite manualmente (40 h/mese, C-14)","Accordi commerciali, contributi, volantini, premi di fine anno","Cliente · anno","Contratti/Excel commerciale","COMM","DISC","Dove sono registrati accordi e contributi trade (ERP, Excel, contratti)? Potete estrarre il trade spend 2024-2025 per insegna?","S"),
 ("K17","RET","X","L3","Forecast accuracy / copertura stock","1 − |forecast − actual| / actual","Lead time Cina + lanci 2026-27","Forecast di vendita vs actual; giacenze","SKU · mese","ERP / Excel pianificazione","AMM","DISC","Come costruite il forecast di vendita e acquisto (ERP, Excel)? Chi lo fa? Ci date forecast vs actual 2025?","C"),

 # ---------------- OEM / WHITE LABEL ----------------
 ("K20","OEM","1","L3","Prospect OEM identificati (ICP) e raggiunti","N. aziende ICP in lista / contattate","Molti OEM non sanno che Laica fa white label (Dolcetta via Campione, C-08)","Lista target, fiere, contatti","Mensile","HubSpot (non usato) / Excel / nessuno","COMM","DISC","Come generate oggi opportunità OEM (fiere, rete, inbound, agenti)? Esiste una lista target? Dove tracciate i contatti OEM?","M"),
 ("K21","OEM","2-3","L3","Lead → opportunità OEM (conversion)","Opportunità aperte / lead qualificati","Nessun modulo sales attivo in HubSpot (C-15)","Storico opportunità/offerte OEM","Opportunità · 2023-2026","Excel commerciale / email","COMM","DISC","Esiste un registro delle offerte/progetti OEM (anche Excel)? Campi: cliente, data primo contatto, valore stimato, stato, data chiusura, motivo perdita","M"),
 ("K22","OEM","4","L1","Win rate OEM e valore medio contratto","Opportunità vinte / chiuse; Σ valore vinto / n.","Ordini 250k–4M€ (C-19)","Da K21 + fatturato OEM (K01)","Anno","Excel + ERP","COMM","DISC","Quante opportunità OEM avete vinto/perso nel 2023-2025 e per quale valore?","M"),
 ("K23","OEM","3-4","L3","Durata ciclo di vendita OEM","Data firma − data primo contatto","Ciclo 1–3 anni (C-19)","Date primo contatto e firma","Opportunità","Excel / email","COMM","DISC","Incluso in K21 (date)","S"),
 ("K24","OEM","5-6","L2","Time-to-first-delivery e qualità (difettosità) OEM","Data prima consegna − firma; pezzi difettosi / consegnati","Personalizzazioni funzionali per OEM (C-04)","Ordini OEM, non conformità","Ordine","ERP + Qualità","AMM","DISC","Registrate non conformità / reclami OEM? In che sistema?","C"),
 ("K25","OEM","7","L1","Retention / espansione clienti OEM (rinnovi, nuovi SKU)","Ricavi OEM t da clienti OEM t-1 / ricavi t-1","Contratti pluriennali","Da K01 (flag OEM)","Anno","ERP JGalileo","AMM","DIR","In JGalileo come si distingue un cliente/ordine OEM (tipo cliente, linea, marchio)? Aggiungere flag all'export K01","M"),

 # ---------------- AMAZON (Seller + Vendor) ----------------
 ("K30","AMZ","1","L3","Impression / sessioni / page view per ASIN e marketplace","Sessioni per ASIN × paese","Fatturato quasi solo Amazon.it (C-10)","Business Report (Seller), Traffic diagnostic (Vendor)","ASIN · settimana · 24 mesi","Seller Central / Vendor Central","AMZ","DIR","Seller Central: Business Report 'Dettaglio pagina per ASIN child' mensile 24 mesi, tutti i marketplace EU. Vendor Central: Retail Analytics Sales+Traffic, stesso periodo","M"),
 ("K31","AMZ","3-4","L3","Conversion rate (Unit Session %) e Buy Box %","Unità ordinate / sessioni; % Buy Box","Performance per paese ignota","Business Report","ASIN · mese","Seller Central","AMZ","DIR","Incluso in K30","M"),
 ("K32","AMZ","4","L1","Vendite Amazon per canale (1P Vendor vs 3P Seller), paese, ASIN","Σ ordered revenue / shipped COGS","Mix Vendor/Seller e peso sul 5% ignoti","Sales report","ASIN · mese · 2024-2026","Seller + Vendor Central","AMZ","DIR","Seller: report 'Transazioni per data' (settlement) 2024→oggi. Vendor: Sales Diagnostic 'Shipped revenue/COGS' per ASIN e marketplace","M"),
 ("K33","AMZ","X","L1","Costo totale canale Amazon % vendite (fee+FBA+storage+resi+Ads+coop Vendor)","Σ costi Amazon / vendite Amazon","18% fee; 20-22% (KOI) vs 38-40% (slide 31/07): da riconciliare (C-11)","Settlement report, fatture Vendor (coop, damage allowance), Ads","Mese","Seller/Vendor Central + ERP","CFO","DIR","Settlement report Seller 2025 completo + fatture/detrazioni Vendor 2025 (co-op, marketing allowance, resi, penali) + fatture Amazon Ads","M"),
 ("K34","AMZ","1-3","L3","ACoS / TACoS / ROAS per campagna","Spesa Ads / vendite attribuite; spesa Ads / vendite totali","AMS 50k€/mese (C-13) = ~600k€/anno: verificare","Report campagne SP/SB/SD","Campagna · settimana · 12-24 mesi","Amazon Ads console","AMZ","DIR","Export Amazon Ads (Seller+Vendor): campagne SP/SB/SD, spesa, clic, vendite 14gg, ACoS, per mese e marketplace 2025→oggi. Chi gestisce le campagne (interno/agenzia)?","M"),
 ("K35","AMZ","6","L2","Tasso resi Amazon e motivi","Unità rese / unità vendute","Resi 4,13% (C-12)","Report resi FBA / Vendor returns","ASIN · mese","Seller/Vendor Central","CSA","DIR","Report 'Resi FBA' 2025 con motivo reso + Vendor Returns","S"),
 ("K36","AMZ","6","L2","Rating/recensioni e Account Health","Rating medio, n. recensioni, ODR","Brand awareness debole fuori Italia","Recensioni, voice of customer","ASIN · mese","Seller Central / Brand Analytics","CSA","DIR","Export Voice of the Customer + rating per ASIN; accesso Brand Analytics (Search Query Performance) se Brand Registry attivo","S"),
 ("K37","AMZ","X","L3","Ore/FTE gestione manuale Amazon (promo, listing, CSV)","Ore mese × costo orario","40 h/mese sales (KOI) / 0,4 FTE (marketing, 31/07) (C-14)","Stima strutturata per attività","Mese","Survey interna","AMZ","DISC","Mappiamo le attività manuali Amazon: chi, quanto tempo/settimana, che strumento (CSV, Excel). Confermate 40 h/mese?","S"),
 ("K38","AMZ","6-7","L3","Registrazioni prodotto / attivazione garanzia da acquirenti Amazon","Registrazioni / unità vendute","Leva Amazon→D2C (packaging, garanzia, QR) (C-21)","Registrazioni garanzia, QR scan, app","Mese","Sito/App/HubSpot?","MKT","DISC","Esiste registrazione prodotto o estensione garanzia? Dove finiscono i dati (sito, app, HubSpot)? Quante registrazioni/anno?","M"),

 # ---------------- D2C / E-COMMERCE ----------------
 ("K40","D2C","1","L3","Traffico sito per sorgente (organico, paid, social, referral, Amazon)","Sessioni per canale","Brand noto per caraffe, non per altri prodotti (C-09)","GA4 o altro analytics","Settimana · 24 mesi","GA4? (da verificare)","MKT","DISC","Che analytics usate (GA4, altro)? Ci date accesso in lettura (viewer) a GA4 e Search Console?","M"),
 ("K41","D2C","1","L1","Share of search / brand search volume vs Brita, Philips","Ricerche brand Laica / Σ ricerche brand categoria","Philips anticipa lanci PFAS di 6 mesi (C-22)","Search Console, Google Trends, Brand Analytics","Mese","Search Console + dati pubblici","MKT","DISC","Accesso Search Console; avete studi di brand awareness/notorietà per paese?","S"),
 ("K42","D2C","2-3","L3","Funnel e-commerce: vista prodotto → carrello → checkout → ordine","Tassi step-to-step","Abbandono 67,1%/47%/37,9%; 16 ordini/mese (C-16)","Eventi e-commerce","Settimana · 24 mesi","GA4 + piattaforma e-commerce","MKT","DISC","La piattaforma è WordPress/WooCommerce o Shopify (dato contrastante)? Export ordini 2023→oggi: data, cliente (email hash), SKU, qty, valore, sconto, spedizione, paese, coupon","M"),
 ("K43","D2C","4","L1","Ordini, AOV, ricavi D2C","Ricavi / ordini","16 ordini/mese (C-16)","Ordini e-commerce","Ordine","Piattaforma e-commerce","MKT","DIR","Incluso in K42","M"),
 ("K44","D2C","1-4","L1","CAC D2C e ROAS paid","Spesa marketing D2C / nuovi clienti","Nessun tracciamento; agenzia MoMA su campagne","Spesa Meta/Google Ads, agenzia","Campagna · mese","Meta/Google Ads, fatture agenzia","MKT","DISC","Quali campagne digital fate oggi (Meta, Google, altri), gestite da chi (interno/MoMA)? Export spesa e risultati 2025; budget marketing 2025 per voce","M"),
 ("K45","D2C","7","L1","Tasso di riacquisto filtri e LTV (12/24 mesi)","Clienti con ≥2 ordini / clienti; Σ margine per cliente","Consumabile ricorrente = leva LTV (150k€ ipotesi, C-18)","Da K42 (id cliente)","Cliente","Piattaforma e-commerce","MKT","DIR","Incluso in K42 (id cliente stabile)","M"),
 ("K46","D2C","7","L3","Abbonamenti / riordino automatico filtri","Clienti in subscription / clienti attivi","Benchmark Brita/Philips subscription","Stato funzionalità","—","Piattaforma e-commerce","MKT","DISC","Avete mai offerto abbonamento o promemoria cambio filtro (app, email)?","C"),

 # ---------------- CRM / DATO CLIENTE ----------------
 ("K50","ALL","X","L3","Base contatti marketable (consenso valido) per paese e segmento","Contatti con consenso / totale","15.000 (KOI) vs 20.000 (31/07) contatti, maggioranza senza consenso (C-15)","Export HubSpot contatti con proprietà consenso, fonte, data creazione","Contatto","HubSpot","IT","DIR","Export HubSpot: tutti i contatti + proprietà (lifecycle stage, fonte originale, data creazione, paese, subscription types/consenso, ultima attività). Indicare licenza HubSpot (hub e tier) e n. utenti","M"),
 ("K51","ALL","X","L3","Uso oggetti CRM (aziende, deal, ticket, pipeline)","N. record e attività per oggetto","CRM usato come rubrica (C-15)","Conteggi per oggetto","—","HubSpot","IT","DIR","Export oggetti Companies, Deals (con pipeline e stage), Tickets se presenti; lista utenti e ultimo login","M"),
 ("K52","ALL","X","L3","Data quality: duplicati, campi obbligatori, match ERP↔CRM↔e-commerce","% record completi, % duplicati, % match","Nessuna integrazione; CSV giornalieri (C-23)","Da K01, K42, K50","—","ERP + HubSpot + e-commerce","IT","DIR","Schema delle integrazioni/script CSV attuali (quali file, frequenza, campi) e chiavi comuni (cod. cliente, email, P.IVA)","S"),
 ("K53","ALL","X","L3","Time-to-market contenuti prodotto (scheda pronta per canale)","Giorni da SKU pronto a scheda online su tutti i canali","Senza PIM: reperire info richiede mesi (C-24); PFAS in ritardo di 6 mesi","Timeline lanci recenti","Lancio","Email / cartelle","MKT","DISC","Per 2-3 lanci recenti (es. filtro PFAS): date di disponibilità prodotto, scheda tecnica, landing, listing Amazon, prima campagna. Dove sono oggi schede, foto e manuali (cartelle, DAM)?","M"),

 # ---------------- CUSTOMER SERVICE ----------------
 ("K60","ALL","6","L2","Volume ticket per canale e per 1.000 unità vendute","Ticket / unità vendute ×1000","6.465 ticket 2025: 4.129 email, 2.228 form, 78 altro (C-17)","Log richieste","Ticket · 2024-2025","Casella email / form sito / (tool?)","CS","DISC","Con quale strumento gestite le richieste (Gmail/Outlook condiviso, Zendesk, altro)? Export completo 2025 (o casella/cartelle .mbox/.pst) + export form web","M"),
 ("K61","ALL","6","L2","Tempo di prima risposta e di risoluzione","Mediana/P90 ore dalla richiesta alla prima risposta / chiusura","3,3 gg email (C-17); benchmark ≤24 h","Timestamp richieste/risposte","Ticket","Da K60","CS","DISC","Incluso in K60 (timestamp); chi ha calcolato il 3,3 gg e come?","M"),
 ("K62","ALL","6","L2","Motivi di contatto (tassonomia) e % automatizzabile","Ticket per causale / totale","80% FAQ ricorsive: 39,5% sottovuoto, 26,4% ricambi sottovuoto, 21% caraffe (C-17)","Testo richieste + causali","Ticket","Da K60","CS","DISC","Esiste già una classificazione? Chi l'ha prodotta? Condividete FAQ, manuali, template risposte in uso","M"),
 ("K63","ALL","6","L2","Soddisfazione (CSAT/NPS) post contatto e post acquisto","Media CSAT; %promotori − %detrattori","Nessuna misura citata","Survey","Mese","—","CS","DISC","Misurate soddisfazione o NPS (anche su Amazon feedback)? Se sì, dove?","C"),
 ("K64","ALL","7","L3","Upsell/cross-sell da CS (ricambi, accessori)","Ordini generati da contatto CS / contatti","Upsell/cross-sell quasi nulli (C-25)","Link ordine↔ticket","Mese","—","CS","DISC","Il CS vende/indirizza a ricambi? Si traccia?","C"),
 ("K65","AMZ","6","L2","Ticket CS Amazon (buyer messages) e tempi","Messaggi / ordini; % risposte <24h","2 persone CS Amazon (KOI)","Buyer-Seller Messages, Customer Service Performance","Mese","Seller Central","CSA","DIR","Seller Central: report Customer Service Performance 12 mesi + volume messaggi","S"),
]

# Mappa stadi
STAGE = {"1":"Awareness","2":"Education","3":"Selection","4":"Commit","5":"Onboarding","6":"Impact","7":"Growth","X":"Trasversale",
         "2-3":"Education→Selection","3-4":"Selection→Commit","5-6":"Onboarding→Impact","6-7":"Impact→Growth","1-3":"Awareness→Selection","1-4":"Awareness→Commit"}
MOTION = {"RET":"B2B Retail/Distribuzione","OEM":"B2B OEM/White label","AMZ":"B2C Amazon","D2C":"B2C D2C e-commerce","ALL":"Trasversale"}
