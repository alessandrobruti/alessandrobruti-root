# WBS — RevOps Assessment · Aurigane S.r.l.

`spec.json` è la specifica ricavata dall'Offerta `260909_Off_Aurigane_RevOps assessment v1.0`
del 9 set 2026 (firmata). Il file Excel si rigenera con la skill `wbs-da-contratto`:

```bash
SK=../skills/wbs-da-contratto
python3 $SK/scripts/bilancia.py spec.json                 # il conto torna?
python3 $SK/scripts/build_wbs.py spec.json --out NOME.xlsx
python3 $SK/scripts/verifica.py /tmp/copia-ricalcolata.xlsx --spec spec.json
```

**Versione 0.9: provvisoria.** I parametri non desunti dal contratto — giornate di back
office, ore e chilometri di viaggio, vitto, squadra della review — sono stime da confermare.
La fase «Review con il cliente» è fuori perimetro contrattuale, aggiunta su richiesta.

Differenza strutturale rispetto a Laica: la Fase 4 ha 3 sessioni tecniche da 2 ore erogate
nella stessa giornata, quindi 3 incontri e **1 trasferta**. È il motivo per cui il
generatore ha la colonna `N. trasferte` separata da `N. incontri`.
