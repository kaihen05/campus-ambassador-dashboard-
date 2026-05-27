import urllib.request, json, sys
sys.stdout.reconfigure(encoding='utf-8')

d = json.loads(urllib.request.urlopen('http://127.0.0.1:8765/api/week_compare', timeout=10).read())
ad = d['ambassadors']

# 看看完整 growthTop 中处于临界区的人
print(f"上周精英: {ad['elitePrevCount']} 人  |  本周精英: {ad['eliteCurrCount']} 人  |  Δ: {ad['eliteDelta']:+d}")
print()

# 抓完整大使列表对比
curr = json.loads(urllib.request.urlopen('http://127.0.0.1:8765/api/data', timeout=10).read())
curr_amb = {a['name']: a for a in curr['ambassadorRank']}

# 上周快照
import pickle, pathlib
snap_path = pathlib.Path('snapshots/2026-W20.pkl')
prev = pickle.load(open(snap_path, 'rb'))
prev_amb = {a['name']: a for a in prev['ambassadorRank']}

print("=== 临界区大使（本周 150~250 简历，看看谁可能马上晋级）===")
print(f"  {'大使':12s} {'本周':>6s} {'上周':>6s} {'净增':>6s} {'本周精英?':>9s} {'上周精英?':>9s}")
border = []
for name, a in curr_amb.items():
    r = a.get('resumes', 0)
    if 150 <= r <= 250:
        p = prev_amb.get(name, {})
        pr = p.get('resumes', 0)
        border.append((name, r, pr, r - pr, r >= 200, pr >= 200))
border.sort(key=lambda x: -x[1])
for n, r, pr, add, ce, pe in border:
    print(f"  {n:12s} {r:>6d} {pr:>6d} {add:>+6d} {'是' if ce else '否':>9s} {'是' if pe else '否':>9s}")
