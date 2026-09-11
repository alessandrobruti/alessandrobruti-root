#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bilancia il numero finale: dice quanto vale oggi il margine e, per ogni leva,
quale valore servirebbe per arrivare al margine obiettivo tenendo fermo il resto.

    python3 bilancia.py spec.json                  # foto del margine + sensibilità
    python3 bilancia.py spec.json --target 0.25     # e cosa serve per arrivare al 25%

Serve nel momento in cui il piano è tecnicamente corretto ma il conto non torna.
Le leve sono sempre le stesse quattro — giornate di back office, persone per
sessione, tariffe, costi di trasferta — e quasi sempre solo due sono davvero
negoziabili. Sapere di quanto va mossa ciascuna evita di limare le stime finché
il file quadra, che è il modo più rapido per consegnare un progetto che non sta
in piedi.

Replica le formule del workbook: se i numeri qui e quelli nel Riepilogo non
coincidono, uno dei due ha un errore.
"""
import argparse
import json
import math

ap = argparse.ArgumentParser(description=__doc__,
                             formatter_class=argparse.RawDescriptionHelpFormatter)
ap.add_argument("spec", help="specifica JSON del progetto")
ap.add_argument("--target", type=float, default=None,
                help="margine obiettivo come frazione (0.25 = 25%%); "
                     "default: parametri.margine_minimo o 0.20")
args = ap.parse_args()

SPEC = json.load(open(args.spec, encoding="utf-8"))
P = SPEC["parametri"]
FASI = SPEC["fasi"]
TARGET = args.target if args.target is not None else P.get("margine_minimo", 0.20)

ORE = P["ore_giornata"]
INV = P.get("investimento") or 0


def conta(pers_override=None, back_scale=None, ore_viaggio=None,
          tf=None, tb=None, tv=None, km=None, ped=None, vitto=None):
    """Ricalcola il costo pieno con una leva spostata. None = valore della spec."""
    tf = P.get("tariffa_front", 0) if tf is None else tf
    tb = P.get("tariffa_back", 0) if tb is None else tb
    tv = P.get("tariffa_viaggio", 0) if tv is None else tv
    hv = P.get("ore_viaggio_trasferta", 0) if ore_viaggio is None else ore_viaggio
    km = P.get("km_trasferta", 0) if km is None else km
    ped = P.get("pedaggi", 0) if ped is None else ped
    vit = P.get("vitto", 0) if vitto is None else vitto
    p_auto = P.get("persone_auto") or 1

    front = viaggio = back = 0.0
    sessioni = pers_sess = 0
    for f in FASI:
        inc = f.get("inc") or 0
        ore = f.get("ore") or 0
        pers = f.get("pers") or 0
        if pers and pers_override and f.get("tipo") == "Cliente":
            pers = pers_override
        front += inc * ore * pers / ORE
        back += (f.get("back") or 0) * (back_scale if back_scale is not None else 1)
        trasf = f.get("trasf") if f.get("trasf") is not None else inc
        if f.get("mod") == "In presenza" and trasf and pers:
            viaggio += trasf * pers * hv / ORE
            sessioni += trasf
            pers_sess += trasf * pers

    auto = math.ceil(pers_sess / p_auto) if pers_sess else 0
    trasf = auto * km * km_aci() + auto * ped + pers_sess * vit
    giornate = front * ORE * tf + back * ORE * tb + viaggio * ORE * tv
    return dict(front=front, back=back, viaggio=viaggio, tot=front + back + viaggio,
                ore=(front + back + viaggio) * ORE, costo_giornate=giornate,
                trasferte=trasf, pieno=giornate + trasf, margine=INV - giornate - trasf,
                sessioni=sessioni, pers_sess=pers_sess, auto=auto)


def km_aci():
    return P.get("aci_eur_km", 0)


def risolvi(leva, lo, hi, tol=1e-4):
    """Cerca il valore della leva che porta il margine al target (bisezione)."""
    obiettivo = INV * TARGET

    def m(v):
        return conta(**{leva: v})["margine"] - obiettivo

    if m(lo) * m(hi) > 0:
        return None
    for _ in range(80):
        mid = (lo + hi) / 2
        if m(lo) * m(mid) <= 0:
            hi = mid
        else:
            lo = mid
        if hi - lo < tol:
            break
    return (lo + hi) / 2


def eur(v):
    return f"{v:,.2f} €".replace(",", "·").replace(".", ",").replace("·", ".")


base = conta()
print()
print(f"  {SPEC['progetto'].get('titolo', 'progetto')}")
print(f"  {'─' * 74}")
print(f"  Giornate            front {base['front']:.2f}   back {base['back']:.2f}   "
      f"viaggio {base['viaggio']:.2f}   →  totale {base['tot']:.2f}  ({base['ore']:.1f} h)")
print(f"  Costo giornate      {eur(base['costo_giornate']):>14}")
print(f"  Costo trasferte     {eur(base['trasferte']):>14}   "
      f"({base['sessioni']} trasferte, {base['pers_sess']} presenze, {base['auto']} viaggi auto)")
print(f"  COSTO PIENO         {eur(base['pieno']):>14}")
print(f"  Investimento        {eur(INV):>14}")
pct = base["margine"] / INV if INV else 0
print(f"  MARGINE             {eur(base['margine']):>14}   {pct*100:.2f}%"
      f"      obiettivo {TARGET*100:.0f}%  →  {eur(INV*TARGET)}")
if INV:
    print(f"  Tariffa media implicita sulle ore erogate: "
          f"{eur(INV/base['ore'] if base['ore'] else 0)}/h")

print()
print(f"  PESO DI OGNI VOCE SUL COSTO PIENO")
print(f"  {'─' * 74}")
for nome, val in (("ore front office", base["front"] * ORE * P.get("tariffa_front", 0)),
                  ("ore back office", base["back"] * ORE * P.get("tariffa_back", 0)),
                  ("ore di viaggio", base["viaggio"] * ORE * P.get("tariffa_viaggio", 0)),
                  ("trasferte (ACI + pedaggi + vitto)", base["trasferte"])):
    quota = val / base["pieno"] * 100 if base["pieno"] else 0
    barra = "█" * int(round(quota / 2.5))
    print(f"  {nome:<34} {eur(val):>13}  {quota:5.1f}%  {barra}")

if base["margine"] >= INV * TARGET:
    print()
    print(f"  Il margine è già sopra l'obiettivo: non serve muovere nulla.")
    raise SystemExit(0)

print()
print(f"  COSA SERVE PER ARRIVARE AL {TARGET*100:.0f}%   (una leva alla volta, il resto fermo)")
print(f"  {'─' * 74}")

LEVE = [
    ("back_scale", "Giornate di back office", 0.0, 1.0,
     lambda v: f"da {base['back']:.2f} a {base['back']*v:.2f} giornate  (−{(1-v)*100:.0f}%)"),
    ("pers_override", "Persone per sessione con il cliente", 1.0, 10.0,
     lambda v: f"{v:.2f} persone al posto di quelle a piano"),
    ("tf", "Tariffa front office", 0.0, 1000.0, lambda v: f"{eur(v)}/h"),
    ("tb", "Tariffa back office", 0.0, 1000.0, lambda v: f"{eur(v)}/h"),
    ("ore_viaggio", "Ore di viaggio per trasferta", 0.0, 24.0, lambda v: f"{v:.2f} h a/r"),
    ("km", "Km per trasferta", 0.0, 2000.0, lambda v: f"{v:.0f} km a/r"),
    ("ped", "Pedaggi per trasferta", 0.0, 500.0, lambda v: f"{eur(v)}"),
    ("vitto", "Vitto per persona", 0.0, 500.0, lambda v: f"{eur(v)}"),
]
for leva, nome, lo, hi, descr in LEVE:
    v = risolvi(leva, lo, hi)
    if v is None:
        print(f"  {nome:<36} non basta da sola")
    else:
        print(f"  {nome:<36} {descr(v)}")

print()
print("  Le leve non sono equivalenti: le giornate di back office sono stime e chi eroga")
print("  ha un'opinione vera su quelle; il numero di persone per sessione è una decisione;")
print("  le tariffe sono in genere date. I costi di trasferta sono gli ultimi a spostare")
print("  qualcosa — se il margine dipende dal vitto, il problema è il prezzo, non il vitto.")
