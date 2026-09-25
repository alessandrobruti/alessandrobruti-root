# Albero KPI Laica: modello Bow Tie × motion

Aggiornato al 25/09/2026. Fonte unica: `build/kpi_data.py`, che genera anche l'Excel e l'agenda. I rimandi `C-xx` puntano a `Registro_Affermazioni.md`.

## Logica

```
L1 BUSINESS   ricavi · margine per canale · crescita · retention/NRR · LTV · CAC · payback ≤24m
   ▲ spiegato da
L2 SERVIZIO   OTIF/fill rate · resi · tempi CS · motivi di contatto · rating · sell-out
   ▲ spiegato da
L3 OPERATIVO  traffico · conversioni per stadio · ACoS/TACoS · pipeline OEM · ore manuali · data quality · time-to-market
```

Stadi Bow Tie: 1 Awareness → 2 Education → 3 Selection → 4 **Commit** → 5 Onboarding → 6 Impact → 7 Growth.
Quattro motion con logiche diverse, da non mescolare:
- **RET**: B2B2C retail e distributori (sell-in);
- **OEM**: white label, ciclo 1-3 anni;
- **AMZ**: Amazon Seller/Vendor, cliente non nostro;
- **D2C**: shop proprio.

## Mappa copertura (n. KPI per motion × stadio)

| Motion | Awareness | Education | Selection | Commit | Onboarding | Impact | Growth | Trasversale |
|---|---|---|---|---|---|---|---|---|
| B2B Retail/Distribuzione | · | · | 2 | 3 | · | 3 | 1 | 2 |
| B2B OEM/White label | 1 | 1 | 2 | 2 | 1 | 1 | 1 | · |
| B2C Amazon | 2 | 1 | 2 | 2 | · | 4 | 1 | 2 |
| B2C D2C e-commerce | 3 | 2 | 2 | 2 | · | · | 2 | · |
| Trasversale | · | · | · | · | · | 4 | 1 | 11 |

L'Onboarding (5) è scoperto quasi ovunque: oggi Laica non sa chi ha comprato né se il prodotto è stato attivato. È il gap strutturale del Bow Tie destro (vedi K38).

## L1: KPI Business

| ID | Motion | Stadio | KPI | Formula | Perché per Laica |
|---|---|---|---|---|---|
| K01 | ALL | Trasversale | **Ricavi netti per canale e motion** | Σ fatturato netto (dopo sconti, premi, NC) per motion RET/OEM/AMZ-V/AMZ-S/D2C/Export | Mix 95/5 B2B/online mai documentato; % OEM ignota (C-01, C-03) |
| K02 | ALL | Trasversale | **Margine lordo (contributo) per canale / famiglia / cliente** | (Ricavo netto − COGS − costi variabili canale) / ricavo netto | Erosione margine Amazon 20-22% vs 38-40% (dato conflittuale, C-11) |
| K03 | ALL | Trasversale | **Crescita ricavi YoY per motion e paese** | (Ricavi t / Ricavi t-1) − 1 | Obiettivo piano triennale: consolidamento EU + espansione FR/ES (C-05) |
| K04 | ALL | Trasversale | **Target piano triennale 2026-2028 (ricavi, mix, paesi, budget)** | Target vs actual | Piano da approvare dalla PLC; nessun target numerico noto (C-06, C-07) |
| K05 | ALL | Trasversale | **Concentrazione clienti (top-10 share, Pareto)** | Ricavi top-10 clienti / ricavi totali | Dipendenza da GDO/retail e pochi OEM (ordini 250k-4M€) |
| K06 | ALL | Trasversale | **Costo del servizio clienti per ordine / per ricavo** | Costo CS (FTE×costo) / n. ordini o ricavi | CS = 1 FTE assorbito da ricorsive (CFO, C-17) |
| K07 | ALL | Trasversale | **Payback / ROI iniziative RevOps (≤24 mesi)** | Beneficio cumulato / investimento | Richiesta esplicita CFO ROI ≤24 mesi (C-20) |
| K11 | RET | Commit | **Sell-in per cliente / insegna / famiglia** | Σ ricavi netti per cliente × famiglia | Rete: 10 agenzie IT, KA diretti Unieuro/MediaWorld/Tesco, distributori (C-30) |
| K15 | RET | Growth | **Retention clienti B2B e NRR** | Ricavi anno t da clienti attivi t-1 / ricavi t-1 degli stessi | Rischio perdita spazio a scaffale vs Brita/Philips |
| K16 | RET | Trasversale | **Trade spend (premi, contributi, promo) % sell-in e ROI promo** | Trade spend / sell-in; Δvolume promo / costo promo | Promozioni gestite manualmente (40 h/mese, C-14) |
| K22 | OEM | Commit | **Win rate OEM e valore medio contratto** | Opportunità vinte / chiuse; Σ valore vinto / n. | Ordini 250k–4M€ (C-19) |
| K25 | OEM | Growth | **Retention / espansione clienti OEM (rinnovi, nuovi SKU)** | Ricavi OEM t da clienti OEM t-1 / ricavi t-1 | Contratti pluriennali |
| K32 | AMZ | Commit | **Vendite Amazon per canale (1P Vendor vs 3P Seller), paese, ASIN** | Σ ordered revenue / shipped COGS | Mix Vendor/Seller e peso sul 5% ignoti |
| K33 | AMZ | Trasversale | **Costo totale canale Amazon % vendite (fee+FBA+storage+resi+Ads+coop Vendor)** | Σ costi Amazon / vendite Amazon | 18% fee; 20-22% (KOI) vs 38-40% (slide 31/07): da riconciliare (C-11) |
| K41 | D2C | Awareness | **Share of search / brand search volume vs Brita, Philips** | Ricerche brand Laica / Σ ricerche brand categoria | Philips anticipa lanci PFAS di 6 mesi (C-22) |
| K43 | D2C | Commit | **Ordini, AOV, ricavi D2C** | Ricavi / ordini | 16 ordini/mese vs “poche centinaia di k€” (C-16, C-28) |
| K44 | D2C | Awareness→Commit | **CAC D2C e ROAS paid** | Spesa marketing D2C / nuovi clienti | Nessun tracciamento; agenzia MoMA su campagne |
| K45 | D2C | Growth | **Tasso di riacquisto filtri e LTV (12/24 mesi)** | Clienti con ≥2 ordini / clienti; Σ margine per cliente | Consumabile ricorrente = leva LTV (150k€ ipotesi, C-18) |

## L2: KPI Servizio

| ID | Motion | Stadio | KPI | Formula | Perché per Laica |
|---|---|---|---|---|---|
| K12 | RET | Impact | **Sell-out e rotazione a scaffale (ove disponibile)** | Pezzi venduti al consumatore / punti vendita / settimane | Senza sell-out non si chiude sell-in↔consumatore (Bruti 21/07) |
| K13 | RET | Impact | **OTIF / fill rate verso retailer** | Righe consegnate complete e puntuali / righe ordinate | Produzione in Cina → rischio stock-out e lead time |
| K14 | RET | Impact | **Resi e NC da trade (% ricavi)** | Valore NC reso / ricavi | Resi 4,13% (riferito ad Amazon? da chiarire, C-12) |
| K24 | OEM | Onboarding→Impact | **Time-to-first-delivery e qualità (difettosità) OEM** | Data prima consegna − firma; pezzi difettosi / consegnati | Personalizzazioni funzionali per OEM (C-04) |
| K35 | AMZ | Impact | **Tasso resi Amazon e motivi** | Unità rese / unità vendute | Resi 4,13% (C-12) |
| K36 | AMZ | Impact | **Rating/recensioni e Account Health** | Rating medio, n. recensioni, ODR | Brand awareness debole fuori Italia |
| K60 | ALL | Impact | **Volume ticket per canale e per 1.000 unità vendute** | Ticket / unità vendute ×1000 | 6.465 ticket 2025: 4.129 email, 2.228 form, 78 altro (C-17) |
| K61 | ALL | Impact | **Tempo di prima risposta e di risoluzione** | Mediana/P90 ore dalla richiesta alla prima risposta / chiusura | 3,3 gg email (C-17); benchmark ≤24 h |
| K62 | ALL | Impact | **Motivi di contatto (tassonomia) e % automatizzabile** | Ticket per causale / totale | 80% FAQ ricorsive: 39,5% sottovuoto, 26,4% ricambi sottovuoto, 21% caraffe (C-17) |
| K63 | ALL | Impact | **Soddisfazione (CSAT/NPS) post contatto e post acquisto** | Media CSAT; %promotori − %detrattori | Nessuna misura citata |
| K65 | AMZ | Impact | **Ticket CS Amazon (buyer messages) e tempi** | Messaggi / ordini; % risposte <24h | 2 persone CS Amazon (KOI) |

## L3: KPI Operativo

| ID | Motion | Stadio | KPI | Formula | Perché per Laica |
|---|---|---|---|---|---|
| K10 | RET | Selection | **Nuovi punti vendita / nuovi clienti B2B attivati** | N. clienti con primo ordine nel periodo | Strategia espansione FR/ES richiede nuovi distributori |
| K17 | RET | Trasversale | **Forecast accuracy / copertura stock** | 1 − |forecast − actual| / actual | Lead time Cina + lanci 2026-27 |
| K18 | RET | Commit | **Ordini per canale di acquisizione (portale B2B, email agente, Vendor)** | N. ordini e valore per canale di ingresso | Ordini da email, portale B2B proprietario, Seller Central manuale (C-31) |
| K19 | RET | Selection→Commit | **Produttività rete agenti/KA (ricavi e nuovi clienti per agenzia)** | Ricavi / agenzia; clienti attivi / agenzia | 5+5 agenzie IT, Pedro ES, distributori (C-30) |
| K20 | OEM | Awareness | **Prospect OEM identificati (ICP) e raggiunti** | N. aziende ICP in lista / contattate | Molti OEM non sanno che Laica fa white label (Dolcetta via Campione, C-08) |
| K21 | OEM | Education→Selection | **Lead → opportunità OEM (conversion)** | Opportunità aperte / lead qualificati | Nessun modulo sales attivo in HubSpot (C-15) |
| K23 | OEM | Selection→Commit | **Durata ciclo di vendita OEM** | Data firma − data primo contatto | Ciclo 1–3 anni (C-19) |
| K30 | AMZ | Awareness | **Impression / sessioni / page view per ASIN e marketplace** | Sessioni per ASIN × paese | Fatturato quasi solo Amazon.it (C-10) |
| K31 | AMZ | Selection→Commit | **Conversion rate (Unit Session %) e Buy Box %** | Unità ordinate / sessioni; % Buy Box | Performance per paese ignota |
| K34 | AMZ | Awareness→Selection | **ACoS / TACoS / ROAS per campagna** | Spesa Ads / vendite attribuite; spesa Ads / vendite totali | AMS 50k€/mese (C-13) = ~600k€/anno: verificare |
| K37 | AMZ | Trasversale | **Ore/FTE gestione manuale Amazon (promo, listing, CSV)** | Ore mese × costo orario | 40 h/mese sales (KOI) / 0,4 FTE (marketing, 31/07) (C-14) |
| K38 | AMZ | Impact→Growth | **Registrazioni prodotto / attivazione garanzia da acquirenti Amazon** | Registrazioni / unità vendute | Leva Amazon→D2C (packaging, garanzia, QR) (C-21) |
| K40 | D2C | Awareness | **Traffico sito per sorgente (organico, paid, social, referral, Amazon)** | Sessioni per canale | Brand noto per caraffe, non per altri prodotti (C-09) |
| K42 | D2C | Education→Selection | **Funnel e-commerce: vista prodotto → carrello → checkout → ordine** | Tassi step-to-step | Abbandono 67,1%/47%/37,9%; 16 ordini/mese (C-16) |
| K46 | D2C | Growth | **Abbonamenti / riordino automatico filtri** | Clienti in subscription / clienti attivi | Benchmark Brita/Philips subscription |
| K50 | ALL | Trasversale | **Base contatti marketable (consenso valido) per paese e segmento** | Contatti con consenso / totale | 20.000 anagrafiche, 15.820 con consenso (C-15): KOI errato |
| K51 | ALL | Trasversale | **Uso oggetti CRM (aziende, deal, ticket, pipeline)** | N. record e attività per oggetto | CRM usato come rubrica (C-15) |
| K52 | ALL | Trasversale | **Data quality: duplicati, campi obbligatori, match ERP↔CRM↔e-commerce** | % record completi, % duplicati, % match | Integrazione ERP↔Seller: live (NZ) vs CSV (SPICED) (C-23) |
| K53 | ALL | Trasversale | **Time-to-market contenuti prodotto (scheda pronta per canale)** | Giorni da SKU pronto a scheda online su tutti i canali | Senza PIM: reperire info richiede mesi (C-24); PFAS in ritardo di 6 mesi |
| K64 | ALL | Growth | **Upsell/cross-sell da CS (ricambi, accessori)** | Ordini generati da contatto CS / contatti | Upsell/cross-sell 1-2% (C-17) |

## Albero causale: i tre KPI che il CFO guarderà

```
Margine contributo per canale (K02)
├── Costo canale Amazon % (K33) ← fee 18% · FBA 20-22% · ADS 595k€ (K34) · resi 4,13% (K35)
├── Mix canale (K01) ← quanto Seller migra su D2C (K43) ← registrazioni prodotto (K38) ← packaging/garanzia
└── Costo servizio (K06) ← ticket/1000 unità (K60) × tempo (K61) × % automatizzabile (K62)

LTV consumatore (K45)
├── Tasso di riacquisto filtri a 12 mesi (K45) ← promemoria/abbonamento (K46) ← consenso marketing (K50)
└── Perdita verso compatibili Brita/Philips (C-49) ← share of search (K41)

Crescita OEM (K22 × K25)
├── Prospect ICP raggiunti (K20) → opportunità (K21) → win rate (K22) → durata ciclo (K23)
└── Retention ed espansione OEM (K25)
```
