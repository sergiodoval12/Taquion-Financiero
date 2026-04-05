import json, re

with open('/sessions/sweet-gracious-mayer/mnt/Taquion Financiero/index.html', 'r') as f:
    content = f.read()

# Simulate the new runway logic
match = re.search(r'const DATA = ({.*?});\s*\n', content)
data = json.loads(match.group(1))
mov = data.get('mov', [])

saldoActual = -65654846
descubierto = 175948095
disponibleConDesc = saldoActual + descubierto

from datetime import datetime, timedelta

def get_week_mon(date_str):
    d = datetime.strptime(date_str, '%Y-%m-%d')
    monday = d - timedelta(days=d.weekday())
    return monday.strftime('%Y-%m-%d')

weeklyAll = {}
for m in mov:
    f = m.get('f','')
    if not f or f < '2026-04-06': continue
    wk = get_week_mon(f)
    if wk not in weeklyAll:
        weeklyAll[wk] = 0
    weeklyAll[wk] += m.get('v', 0) or 0

weekKeys = sorted(weeklyAll.keys())
runCash = saldoActual
criticalWeek = None
minConDesc = disponibleConDesc
minWeek = None

print("=== NEW RUNWAY PROJECTION (weekly, ALL movements) ===")
for i, wk in enumerate(weekKeys):
    runCash += weeklyAll[wk]
    conDesc = runCash + descubierto
    end = (datetime.strptime(wk, '%Y-%m-%d') + timedelta(days=6)).strftime('%m/%d')
    start = datetime.strptime(wk, '%Y-%m-%d').strftime('%m/%d')
    marker = ""
    if conDesc < minConDesc:
        minConDesc = conDesc
        minWeek = wk
    if conDesc < 0 and not criticalWeek:
        criticalWeek = wk
        marker = " ⚠️ RUNWAY ENDS HERE"
    print(f"  {start}-{end}: ConDesc={conDesc/1000:>10,.0f}K{marker}")

if criticalWeek:
    weeks_to_critical = weekKeys.index(criticalWeek) + 1
    months = weeks_to_critical / 4.33
    print(f"\n  🔴 Runway: {months:.1f} meses ({weeks_to_critical} semanas)")
    print(f"  Se queda sin plata en semana del {criticalWeek}")
else:
    print(f"\n  🟢 No se queda sin plata en {len(weekKeys)} semanas ({len(weekKeys)/4.33:.0f} meses)")
    print(f"  Mínimo ConDesc: {minConDesc/1000:,.0f}K en semana {minWeek}")
