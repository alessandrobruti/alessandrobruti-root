# wbs-da-contratto

Skill che da un'offerta tecnico-economica produce una WBS Excel a cinque fogli con
formule vive, Gantt settimanale, costi e margine.

Sviluppata sul progetto Laica S.p.A. (`wbs-assessment/` in questo repo):
`scripts/spec.esempio.json` è la specifica di quel progetto e il generatore la
riproduce identica sui valori chiave, così la skill ha un test di regressione reale.

```bash
python3 scripts/build_wbs.py scripts/spec.esempio.json --out prova.xlsx
cp prova.xlsx /tmp/v.xlsx && python3 <xlsx-skill>/scripts/recalc.py /tmp/v.xlsx 300
python3 scripts/verifica.py /tmp/v.xlsx --spec scripts/spec.esempio.json
```
