# -*- coding: utf-8 -*-
"""
通用基线设置工具 — 把任意一个历史 Excel 跑成「上周基线」(snapshot pkl)。

用法：
  # 1) 用文件名里包含的日期(如 5.14_最终版.xlsx) 自动算出 ISO 周(W20)
  python set_baseline.py "/path/to/5.14_最终版.xlsx"

  # 2) 显式指定周键
  python set_baseline.py "/path/to/xxx.xlsx" --week 2026-W20

  # 3) 用文件「最后修改时间」作为周键
  python set_baseline.py "/path/to/xxx.xlsx" --use-mtime

  # 4) 强制覆盖已存在的同名快照
  python set_baseline.py "..." --force

机制说明：
  - /api/week_compare 默认会自动选择"非本周"中最新的快照作为基线
  - 所以你只要把"上周文件"喂进来打一份基线，本周看板就能自动 vs 上周
  - 跑完后强刷看板即可（无需重启 server）
"""
import argparse
import datetime as dt
import io
import os
import pickle
import re
import sys
import time

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import server  # noqa: E402

SNAPSHOT_DIR = os.path.join(HERE, 'snapshots')
os.makedirs(SNAPSHOT_DIR, exist_ok=True)


# ------------------------------------------------------------------
# 周键推断
# ------------------------------------------------------------------
def _iso_week(d: dt.datetime) -> str:
    y, w, _ = d.isocalendar()
    return f"{y}-W{w:02d}"


def _guess_week_key(file_path: str, use_mtime: bool) -> str:
    """优先级：--use-mtime > 文件名里 5.14 / 2026-05-14 等日期 > 文件 mtime。"""
    name = os.path.basename(file_path)
    if not use_mtime:
        # 5.14 / 5_14 / 5-14 风格 → 当年/月/日
        m = re.search(r'(\d{1,2})[._-](\d{1,2})', name)
        if m:
            mo, da = int(m.group(1)), int(m.group(2))
            year = dt.datetime.now().year
            try:
                return _iso_week(dt.datetime(year, mo, da))
            except ValueError:
                pass
        # 2026-05-14 / 20260514 风格
        m = re.search(r'(20\d{2})[-_]?(\d{2})[-_]?(\d{2})', name)
        if m:
            try:
                return _iso_week(dt.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))))
            except ValueError:
                pass
    # fallback: mtime
    return _iso_week(dt.datetime.fromtimestamp(os.path.getmtime(file_path)))


# ------------------------------------------------------------------
# 选 sheet
# ------------------------------------------------------------------
PREFERRED_SHEETS = (
    '获取名单上的这些外部伯乐的简历推荐数据_1',
    '获取名单上的这些外部伯乐的简历推荐数据',
    'Sheet0',
)


def _pick_sheet(file_bytes: bytes) -> str:
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    for s in PREFERRED_SHEETS:
        if s in xl.sheet_names:
            return s
    # 关键词命中
    for s in xl.sheet_names:
        if any(k in s for k in ('外部伯乐', '简历推荐', '伯乐')):
            return s
    # 行数最多的
    best, best_rows = xl.sheet_names[0], 0
    for s in xl.sheet_names:
        try:
            n = len(pd.read_excel(io.BytesIO(file_bytes), sheet_name=s))
            if n > best_rows:
                best_rows, best = n, s
        except Exception:
            pass
    return best


# ------------------------------------------------------------------
# 主流程
# ------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('excel', help='Excel 文件路径')
    ap.add_argument('--week', help="ISO 周键，如 2026-W20。不传则自动推断")
    ap.add_argument('--use-mtime', action='store_true', help='强制用文件 mtime 推断周键')
    ap.add_argument('--force', action='store_true', help='覆盖已存在的同名快照')
    args = ap.parse_args()

    excel_path = os.path.abspath(args.excel)
    if not os.path.exists(excel_path):
        print(f"❌ 文件不存在：{excel_path}")
        sys.exit(1)

    week_key = args.week or _guess_week_key(excel_path, args.use_mtime)
    target = os.path.join(SNAPSHOT_DIR, f"{week_key}.pkl")
    print(f"📂 输入文件：{excel_path}")
    print(f"📅 基线周键：{week_key}  →  {target}")
    if os.path.exists(target) and not args.force:
        print(f"⚠️  快照已存在；如需覆盖请加 --force")
        sys.exit(2)

    with open(excel_path, 'rb') as f:
        file_bytes = f.read()
    sheet = _pick_sheet(file_bytes)
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
    print(f"  📑 选中 sheet: {sheet}  ({len(df):,} 行)")

    server.load_category_map()
    df_uhr = server.load_uhr_df()
    print(f"⚙️  跑 process_data ...")
    t0 = time.time()
    data = server.process_data(df, df_uhr)
    print(f"  完成，耗时 {time.time() - t0:.1f}s")

    s = data['summary']
    print(f"📊 基线核心指标：")
    print(f"   入库简历 = {s['totalResumes']:,}")
    print(f"   Offer    = {s['totalOffers']}")
    print(f"   精英大使 = {s['eliteCount']}")
    print(f"   总Offer率 = {s['offerRate']}%")
    print(f"   青云大使 = {s['qingyunResumes']} 简历 / {s['qingyunOffers']} Offer")

    snap = {
        'week_key': week_key,
        'snapshot_time': time.time(),
        'summary': data['summary'],
        'ambassadorRank': data.get('ambassadorRank'),
        'regionRank': data.get('regionRank'),
        'uhrRank': data.get('uhrRank'),
        'source_file': os.path.basename(excel_path),
    }
    with open(target, 'wb') as f:
        pickle.dump(snap, f)
    print(f"💾 已写入 {target}")
    print(f"")
    print(f"✅ 下一步：")
    print(f"   1) 本地：浏览器强刷 http://127.0.0.1:8765/  → 看「📈 本周新增 — vs 上周基线」")
    print(f"   2) 远端：把这个 pkl 推到部署服务器：")
    print(f"      scp '{target}' user@host:/path/to/snapshots/{week_key}.pkl")


if __name__ == '__main__':
    main()
