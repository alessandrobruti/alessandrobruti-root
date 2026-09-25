# -*- coding: utf-8 -*-
"""Grafici del report Laica — palette Impresoft: giallo #FDC300 = serie in evidenza, grigi puri per il resto."""
import os, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Polygon, Rectangle
from matplotlib import font_manager as fm

for f in os.listdir(os.path.expanduser("~/.fonts")):
    if f.endswith(".ttf"): fm.fontManager.addfont(os.path.join(os.path.expanduser("~/.fonts"), f))
Y, K, G1, G2, G3, G4, RULE = "#FDC300", "#000000", "#3A3A3A", "#6E6E6E", "#9B9B9B", "#C9C9C9", "#DCDCDC"
plt.rcParams.update({"font.family": "Manrope", "font.size": 10, "axes.edgecolor": RULE, "axes.labelcolor": K,
    "xtick.color": G1, "ytick.color": G1, "axes.spines.top": False, "axes.spines.right": False,
    "axes.titleweight": "bold", "axes.titlesize": 12, "axes.titlelocation": "left", "figure.dpi": 150})
OUT = os.path.join(os.path.dirname(__file__), "..", "Report", "img"); os.makedirs(OUT, exist_ok=True)
def save(fig, name):
    fig.savefig(os.path.join(OUT, name), bbox_inches="tight", facecolor="white"); plt.close(fig)
def note(ax, t, y=-0.18):
    ax.text(0, y, t, transform=ax.transAxes, fontsize=7.5, color=G2, va="top")

# 1. Mix ricavi: due definizioni del 95/5
def mix():
    fig, axes = plt.subplots(2, 1, figsize=(8.5, 2.9))
    rows = [("KOI 25/09 (Campione)", [("B2B retail + OEM", 95, K, "white"), ("Online: Amazon Seller+Vendor + shop", 5, Y, K)]),
            ("Call 03/08 (cliente)", [("OEM + GDO/retail + Amazon VENDOR", 95, K, "white"), ("Seller + shop D2C", 5, Y, K)])]
    for ax, (t, parts) in zip(axes, rows):
        left = 0
        for lab, v, c, tc in parts:
            ax.barh(0, v, left=left, color=c, height=0.6)
            ax.text(left + v/2 if v > 10 else left + v + 0.8, 0, f"{lab}  {v}%", ha="center" if v > 10 else "left", va="center", color=tc if v > 10 else K, fontsize=8.5, fontweight="bold")
            left += v
        ax.set_xlim(0, 135); ax.axis("off"); ax.text(0, 0.55, t, fontsize=8.5, color=G1)
    axes[0].set_title("Mix ricavi 95/5: due definizioni diverse di cosa sta nel 5%")
    note(axes[1], "Su ricavi Laica ~€25m (bilancio 2024, conf. C) il 5% vale ~€1,25m. Quota OEM: IGNOTA. Amazon Vendor: B2B o online? da chiarire (K01).", -0.25)
    save(fig, "01_mix_ricavi.png")

# 2. Customer service ticket
def cs():
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(9, 3.2), gridspec_kw={"width_ratios": [1, 1.2]})
    ch = [("Email", 4129, "3,3 gg"), ("Form web", 2228, "9,1 gg"), ("Facebook", 78, "n.d.")]
    bars = a1.barh([c[0] for c in ch][::-1], [c[1] for c in ch][::-1], color=[G4, G3, Y])
    for b, (_, v, t) in zip(bars, ch[::-1]): a1.text(b.get_width() + 60, b.get_y() + b.get_height()/2, f"{v:,}".replace(",", ".") + f"  · 1ª risp. {t}", va="center", fontsize=8.5)
    a1.set_title("Ticket 2025: 6.435 per canale"); a1.set_xlim(0, 6500); a1.xaxis.set_visible(False); a1.spines["bottom"].set_visible(False)
    note(a1, "Fonte: risposte scritte cliente (Zanuso 02/09) via SPICED, conf. B.\nKOI riportava 6.465 e “9,9 gg”: errori di lettura.", -0.08)
    mot = [("Assistenza sottovuoto", 39.5), ("Ricambi sottovuoto", 26.4), ("Caraffe", 21.0), ("Altro / non class.", 13.1)]
    bars = a2.barh([m[0] for m in mot][::-1], [m[1] for m in mot][::-1], color=G3)
    bars[-1].set_color(Y)
    for b in bars: a2.text(b.get_width() + 0.8, b.get_y() + b.get_height()/2, f"{b.get_width():.1f}%", va="center", fontsize=9)
    a2.set_title("Motivi di contatto (quota ticket)"); a2.set_xlim(0, 50); a2.xaxis.set_visible(False); a2.spines["bottom"].set_visible(False)
    note(a2, "“80% FAQ ricorsive” ma 3 motivi = 86,9%. Sottovuoto = 65,9%:\nil driver del CS è il sottovuoto, non la filtrazione.", -0.08)
    fig.suptitle("Customer service: 1 persona, nessun tool, nessuna automazione", x=0.01, ha="left", fontweight="bold", fontsize=12)
    fig.tight_layout(); save(fig, "02_customer_service.png")

# 3. Funnel e-commerce
def funnel():
    fig, ax = plt.subplots(figsize=(8, 3))
    steps = [("Carrello", 100), ("Checkout", 53.0), ("Acquisto", 32.9)]
    for i, (s, v) in enumerate(steps):
        ax.barh(i, v, color=Y if i == 2 else G4, height=0.6, left=(100 - v)/2)
        ax.text(50, i, f"{s}  {v:.0f}%", ha="center", va="center", fontweight="bold", fontsize=10)
    ax.invert_yaxis(); ax.axis("off")
    ax.set_title("Shop ricambi (WooCommerce): funnel carrello→ordine")
    ax.text(102, 0.5, "−47% carrello→checkout", fontsize=8.5, color=G1, va="center")
    ax.text(102, 1.5, "−37,9% checkout→acquisto", fontsize=8.5, color=G1, va="center")
    ax.set_xlim(0, 140)
    note(ax, "Abbandono cumulato 67,1% = 1−0,53×0,621 → coerente. 16 ordini/mese ≈ 190/anno ≈ pochi k€.\nContraddizione: 03/08 “D2C = poche centinaia di migliaia €”. “0,08% sul database” = 16/20.000: non è una conversione.", -0.02)
    save(fig, "03_funnel_ecommerce.png")

# 4. Amazon: costo canale
def amz():
    fig, ax = plt.subplots(figsize=(8.8, 3.3))
    rows = [("Cliente 15/07 + risposte scritte", [18, 21, 0]), ("+ ADS 595k€/anno (quota ignota)", [18, 21, 10]), ("KOI 25/09 (lettura errata)", [0, 21, 0])]
    cols = [Y, G3, G1]; lab = ["Referral fee 18%", "FBA/storage/resi 20-22%", "ADS (ipotesi TACoS 10%)"]
    for i, (r, v) in enumerate(rows):
        left = 0
        for j, x in enumerate(v):
            if x == 0: continue
            ax.barh(i, x, left=left, color=cols[j], height=0.55)
            ax.text(left + x/2, i, f"{x}%" if j < 2 else "~10%?", ha="center", va="center", fontsize=8.5, color=K if j == 0 else "white")
            left += x
        ax.text(left + 1, i, {0: "38–40% (“>40%” nello SPICED)", 1: "~48–50%?", 2: "20–22% come totale"}[i], va="center", fontweight="bold", fontsize=9)
    ax.set_yticks(range(3)); ax.set_yticklabels([r[0] for r in rows]); ax.invert_yaxis()
    ax.set_xlim(0, 75); ax.xaxis.set_visible(False); ax.spines["bottom"].set_visible(False)
    ax.set_title("Costo Amazon Seller % vendite: cosa sappiamo e cosa no")
    note(ax, "Referral fee pubblica Amazon.it Casa e cucina 15% (8% sotto 20€): il 18% dichiarato va riconciliato (mix categorie, DST 3%?). ADS: nessun dato di ACoS/TACoS.\nFonti: meeting report 15/07 (Zanuso, B) · SPICED (B) · KOI (C) · tariffario Amazon (web, C).", -0.05)
    save(fig, "04_amazon_costo.png")

# 5. Ipotesi di valore: robustezza
def ipotesi():
    fig, ax = plt.subplots(figsize=(8.8, 3.6))
    H = [("Riattivazione DB: 20.000×25%×30€ (slide)", 150, "D"), ("…con 15.820 contatti con consenso", 118.6, "D"),
         ("…25% è un'ipotesi: con 5% di riacquisto", 23.7, "D"),
         ("Efficienza FTE (presentazione): mkt 0,25 + CS 0,7", 43.7, "C"), ("Recupero fee 18% su 10% del Seller (ipotesi)", 0, "E")]
    y = range(len(H))
    bars = ax.barh(list(y), [h[1] for h in H], color=[G3 if h[2] == "D" else Y for h in H], height=0.6)
    for b, h in zip(bars, H): ax.text(b.get_width() + 2, b.get_y() + b.get_height()/2, (f"{h[1]:.0f}k€/anno · conf. {h[2]}" if h[1] else "non calcolabile: fatturato Seller ignoto"), va="center", fontsize=8.5)
    ax.set_yticks(list(y)); ax.set_yticklabels([h[0] for h in H]); ax.invert_yaxis()
    ax.set_xlim(0, 230); ax.xaxis.set_visible(False); ax.spines["bottom"].set_visible(False)
    ax.set_title("Business case pre-vendita: quanto regge ogni ipotesi")
    note(ax, "Giallo = input dichiarato da owner (40% tempo Bellero; 1 FTE CS da Zanuso) · grigio = moltiplicatori Impresoft. Nessun tasso di riacquisto misurato: è il primo dato da chiedere (K45).", -0.05)
    save(fig, "05_ipotesi_valore.png")

# 9. Ricavi Laica e Strix CG
def ricavi():
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(9, 3))
    x = ["2019", "2023", "2024"]; v = [20.4, 24.34, 25.15]
    b = a1.bar(x, v, color=[G4, G3, Y], width=0.6)
    for bb, vv in zip(b, v): a1.text(bb.get_x() + bb.get_width()/2, vv + 0.4, f"€{vv:.1f}m", ha="center", fontsize=9)
    a1.set_ylim(0, 30); a1.set_title("Laica S.p.A. ricavi"); a1.yaxis.set_visible(False); a1.spines["left"].set_visible(False)
    note(a1, "2019: RNS acquisizione Strix (A). 2023-24: aggregatori\nbilancio (C). Utile 2024 €1,6m. 2020-22 n.d.", -0.12)
    x2 = ["FY22", "FY23", "FY24", "12m a\n03/2026"]; v2 = [35.4, 32.3, 31.4, 34.4]
    b2 = a2.bar(x2, v2, color=[G4, G4, G3, Y], width=0.6)
    for bb, vv in zip(b2, v2): a2.text(bb.get_x() + bb.get_width()/2, vv + 0.5, f"£{vv:.1f}m", ha="center", fontsize=9)
    a2.set_ylim(0, 42); a2.set_title("Strix Consumer Goods (Laica+Aqua Optima+OEM)"); a2.yaxis.set_visible(False); a2.spines["left"].set_visible(False)
    note(a2, "Fonte: RNS Strix (A, via estratti). Crescita 2025-26 trainata da OEM\n(baby brand, filtri bespoke), non dai brand.", -0.12)
    fig.tight_layout(); save(fig, "09_ricavi.png")

# 10. Competitor per dimensione
def comp():
    fig, ax = plt.subplots(figsize=(8.5, 3.4))
    C = [("BWT", 1025, "filtrazione"), ("Omron Healthcare", 900, "health"), ("Brita (gruppo)", 691, "filtrazione"),
         ("Beurer", 308, "health/bilance"), ("Tenacta/Imetec", 126, "SDA/beauty"), ("Brita Italia", 32, "filtrazione"), ("Laica", 25, "ibrido")]
    bars = ax.barh([c[0] for c in C][::-1], [c[1] for c in C][::-1], color=[Y if c[0] == "Laica" else G3 for c in C][::-1])
    for b, c in zip(bars, C[::-1]): ax.text(b.get_width() * 1.08, b.get_y() + b.get_height()/2, f"€{c[1]}m · {c[2]}", va="center", fontsize=8.5)
    ax.set_xscale("log"); ax.set_xlim(10, 6000); ax.xaxis.set_visible(False); ax.spines["bottom"].set_visible(False)
    ax.set_title("Dimensione competitor (ricavi, scala log)")
    note(ax, "Omron ¥145bn convertito ~€0,9bn. Philips water = licenziatari (n.d.). Fonti web A/C, anni 2023-2026. Laica compete con player 10-40× più grandi.", -0.05)
    save(fig, "10_competitor.png")

# 11. Affidabilità delle fonti
def fonti():
    fig, ax = plt.subplots(figsize=(8.5, 2.6))
    lv = [("A documentale", 11), ("B dichiarato da owner", 27), ("C riportato", 6), ("D stima IE", 5), ("E conflitto/ignoto", 9)]
    left = 0
    for i, (l, v) in enumerate(lv):
        ax.barh(0, v, left=left, color=[K, Y, G3, G4, "#F4F4F4"][i], height=0.5, edgecolor=RULE)
        ax.text(left + v/2, 0, str(v), ha="center", va="center", color="white" if i == 0 else K, fontweight="bold")
        ax.text(left + v/2, -0.45, l, ha="center", fontsize=7.5)
        left += v
    ax.set_xlim(0, left); ax.set_ylim(-0.7, 0.4); ax.axis("off")
    ax.set_title("Registro affermazioni: 58 claim per livello di confidenza")
    save(fig, "11_fonti.png")

# 6. Bow tie
def bowtie():
    fig, ax = plt.subplots(figsize=(10, 4.4)); ax.set_xlim(0, 14); ax.set_ylim(-0.4, 6.2); ax.axis("off")
    st = ["Awareness", "Education", "Selection", "Commit", "Onboarding", "Impact", "Growth"]
    xs = [1, 3, 5, 7, 9, 11, 13]; hs = [5, 4, 3, 1.6, 3, 4, 5]
    ax.add_patch(Polygon([[0, 5.3], [7, 3.8], [7, 2.2], [0, 0.7]], closed=True, fc="#F4F4F4", ec=RULE))
    ax.add_patch(Polygon([[14, 5.3], [7, 3.8], [7, 2.2], [14, 0.7]], closed=True, fc="#F4F4F4", ec=RULE))
    gap = {0: "3", 1: "2", 2: "2", 3: "2", 4: "3", 5: "3", 6: "3"}  # 3 = assente, 2 = parziale
    status = {"Awareness": "Brand debole fuori IT;\nOEM non sanno", "Education": "Contenuti lenti\n(no PIM): PFAS −6m",
              "Selection": "Amazon .it only;\nshop solo ricambi", "Commit": "16 ordini/mese D2C;\nOEM 1-3 anni",
              "Onboarding": "Nessuna registraz.\nprodotto tracciata", "Impact": "CS 3,3 gg; resi 4,1%;\nno CSAT",
              "Growth": "Riacquisto filtri\nnon presidiato"}
    for i, (s, x) in enumerate(zip(st, xs)):
        ax.add_patch(FancyBboxPatch((x - 0.9, 5.5), 1.8, 0.55, boxstyle="square,pad=0", fc=K))
        ax.text(x, 5.78, s, ha="center", va="center", color="white", fontsize=8.5, fontweight="bold")
        ax.text(x, 3.0, status[s], ha="center", va="center", fontsize=7.2, color=K)
        ax.add_patch(Rectangle((x - 0.45, 1.05), 0.9, 0.25, fc=Y if gap[i] == "3" else G4))
    ax.text(7, 0.2, "Acquisizione  <  Commit  >  Espansione/ricorrenza", ha="center", fontsize=8, color=G2)
    ax.text(0, -0.3, "Barra gialla = stadio non misurato né presidiato · grigia = misurato parzialmente/manuale. Fonte: sintesi call interne (conf. C).", fontsize=7.5, color=G2)
    ax.set_title("Bow Tie Laica B2C: dove si perde valore (as-is)", loc="left", fontweight="bold")
    save(fig, "06_bowtie_asis.png")

# 7. Timeline progetto
def timeline():
    import datetime as dt
    ev = [("Kick-off interno", "2026-09-25"), ("Setup Monday", "2026-09-30"), ("Kick-off cliente", "2026-10-02"),
          ("Deadline dati", "2026-10-09"), ("Analisi dati (Vicenza)", "2026-10-15"), ("Review interna", "2026-10-16"),
          ("Workshop da sett.", "2026-10-19"), ("Roadmap target", "2026-12-04")]
    d = [dt.date.fromisoformat(e[1]) for e in ev]
    fig, ax = plt.subplots(figsize=(10, 2.3))
    ax.axhline(0, color=RULE, lw=2)
    ax.axvspan(dt.date(2026, 10, 5), dt.date(2026, 10, 9), color="#F4F4F4"); ax.text(dt.date(2026, 10, 7), -0.9, "cliente in ferie", ha="center", fontsize=7, color=G2)
    for i, ((n, _), x) in enumerate(zip(ev, d)):
        hi = n in ("Kick-off cliente", "Deadline dati")
        ax.plot(x, 0, "s", color=Y if hi else K, ms=9 if hi else 6)
        ax.text(x, 0.35 if i % 2 == 0 else -0.55, f"{n}\n{x.strftime('%d/%m')}", ha="center", fontsize=7.5, fontweight="bold" if hi else "normal")
    ax.set_ylim(-1.1, 1.1); ax.axis("off")
    ax.set_title("Calendario fase raccolta dati (da KOI 25/09)", loc="left", fontweight="bold")
    save(fig, "07_timeline.png")

# 8. Stakeholder map
def stake():
    fig, ax = plt.subplots(figsize=(8, 5))
    P = [("Dolcetta (CEO CG)", 9.3, 8.5, "Decisore · sponsor"), ("Zanuso (CFO, Strix)", 9.3, 7.3, "Decisore · vuole numeri, ROI≤24m"),
         ("Milich (Sales Dir.)", 6.5, 5.8, "Influencer · vuole D2C subito;\niniz. difensivo"), ("Bellero (Marketing)", 7, 7.8, "Influencer · collaborativa,\ndebole su digital"),
         ("Sella (IT)", 4.5, 3.0, "Operativo · da ingaggiare"), ("Zen (CS)", 3, 5.3, "Operativo · sovraccarica"),
         ("Contin (Amazon)", 3.8, 6.4, "Operativo · Vendor"), ("CS Amazon (2)", 2.2, 3.8, "Operativo"), ("Amministrazione (2)", 5.2, 1.9, "Owner dati ERP")]
    for n, x, y, r in P:
        dec = "Decisore" in r
        ax.scatter(x, y, s=260, color=Y if dec else ("#000" if "Influencer" in r else G4), zorder=3, edgecolor=K, lw=0.6)
        ax.text(x + 0.3, y, f"{n}\n{r}", va="center", fontsize=7.3)
    ax.set_xlim(0, 13.5); ax.set_ylim(0, 10); ax.set_xlabel("Potere decisionale →"); ax.set_ylabel("Interesse/coinvolgimento nel progetto →")
    ax.set_xticks([]); ax.set_yticks([])
    ax.set_title("Mappa stakeholder (posizioni stimate)")
    note(ax, "Posizioni stimate da meeting report IE e call interne (conf. C/D). Registrazioni call cliente non accessibili.", -0.08)
    save(fig, "08_stakeholder.png")

if __name__ == "__main__":
    for f in (mix, cs, funnel, amz, ipotesi, bowtie, timeline, stake, ricavi, comp, fonti): f()
    print(sorted(os.listdir(OUT)))
