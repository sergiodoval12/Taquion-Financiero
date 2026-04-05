import json, re
from datetime import datetime, timedelta

with open('/sessions/sweet-gracious-mayer/mnt/Taquion Financiero/index.html', 'r') as f:
    content = f.read()

match = re.search(r'const DATA = ({.*?});\s*\n', content)
data = json.loads(match.group(1))
mov = data.get('mov', [])

saldoActual = -65654846
descubierto = 175948095

def get_week_mon(ds):
    d = datetime.strptime(ds, '%Y-%m-%d')
    return (d - timedelta(days=d.weekday())).strftime('%Y-%m-%d')

def mov_class(m):
    t = m.get('t','')
    it = (m.get('i','') or '').lower()
    if t == 'Potenciales': return 'potencial'
    if it.startswith('aporte de capital') or it.startswith('préstamo gc') or 'aporte mutuo' in it or 'aporte muto' in it: return 'inversion'
    return 'operativo'

weeklyOp = {}
weeklyFull = {}
for m in mov:
    if not m.get('f') or m['f'] < '2026-04-06': continue
    wk = get_week_mon(m['f'])
    weeklyOp.setdefault(wk, 0)
    weeklyFull.setdefault(wk, 0)
    v = m.get('v',0) or 0
    cls = mov_class(m)
    if cls == 'operativo':
        weeklyOp[wk] += v
        weeklyFull[wk] += v
    elif cls == 'potencial':
        weeklyFull[wk] += v

wks = sorted(set(list(weeklyOp.keys()) + list(weeklyFull.keys())))

cashOp = saldoActual
cashFull = saldoActual

print(f"{'Semana':<14} {'Op ConDesc':>12} {'Full ConDesc':>12}")
for wk in wks[:12]:
    cashOp += weeklyOp.get(wk, 0)
    cashFull += weeklyFull.get(wk, 0)
    cdOp = cashOp + descubierto
    cdFull = cashFull + descubierto
    end = (datetime.strptime(wk, '%Y-%m-%d') + timedelta(days=6)).strftime('%m/%d')
    start = datetime.strptime(wk, '%Y-%m-%d').strftime('%m/%d')
    mOp = " ⚠️" if cdOp < 0 else ""
    mFull = " ⚠️" if cdFull < 0 else ""
    print(f"  {start}-{end}  {cdOp/1000:>10,.0f}K{mOp}  {cdFull/1000:>10,.0f}K{mFull}")

print(f"\nUser Excel: semana 25/5 → -11MM ConDesc")
