"""
Excel 文件夹监听器
监控指定文件夹，新 Excel 文件进入时自动触发数据处理并更新看板
"""
import os
import time
import glob
import shutil
import io
import sys

# 工作目录
WORK_DIR = os.path.dirname(os.path.abspath(__file__))
WATCH_DIR = os.environ.get('WATCH_DIR', os.path.expanduser('~/Desktop'))  # 监控目录（可改环境变量）
TARGET_PATTERN = '*.xlsx'                         # 检测目标
PROCESSED_DIR = os.path.join(WORK_DIR, 'processed')  # 已处理文件归档

# 跳过文件名关键词（临时文件、~$开头）
SKIP_KEYWORDS = ['~$', '.tmp', '.crdownload', '.part']

# 轮询间隔（秒）
POLL_INTERVAL = 5

# 上一次处理的文件（防止重复）
_last_processed = None

def should_process(filename):
    """判断文件是否需要处理"""
    name = os.path.basename(filename)
    for kw in SKIP_KEYWORDS:
        if kw in name:
            return False
    return True

def get_latest_xlsx():
    """获取监控目录下最新的 xlsx 文件（排除临时文件）"""
    pattern = os.path.join(WATCH_DIR, TARGET_PATTERN)
    files = glob.glob(pattern)
    # 过滤
    files = [f for f in files if should_process(f)]
    if not files:
        return None
    # 按修改时间排序，返回最新的
    files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return files[0]

def process_file(file_path):
    """处理单个 Excel 文件，触发 server.py 中的处理逻辑"""
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    import pandas as pd
    import json
    from server import process_data, UHR_PATH

    print(f"\n🆕 检测到新文件: {os.path.basename(file_path)}")

    try:
        with open(file_path, 'rb') as f:
            file_bytes = f.read()

        # 找目标 sheet
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        target_sheet = None
        for name in xl.sheet_names:
            if '外部伯乐' in name or '简历推荐' in name:
                target_sheet = name
                break
        if not target_sheet:
            target_sheet = xl.sheet_names[0]

        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet)
        print(f"  📊 读取到 {len(df)} 条记录，sheet='{target_sheet}'")

        with open(UHR_PATH, 'rb') as uf:
            df_uhr = pd.read_excel(io.BytesIO(uf.read()), sheet_name='Sheet2')

        data = process_data(df, df_uhr)

        # 保存 JSON
        json_path = os.path.join(WORK_DIR, 'dashboard_data.json')
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        # 归档文件
        os.makedirs(PROCESSED_DIR, exist_ok=True)
        ts = time.strftime('%Y%m%d_%H%M%S')
        archive_name = f"{ts}_{os.path.basename(file_path)}"
        archive_path = os.path.join(PROCESSED_DIR, archive_name)
        shutil.move(file_path, archive_path)
        print(f"  ✅ 处理完成，JSON已更新，文件归档至: {archive_name}")
        print(f"  📊 当前看板数据：{data['summary']['totalResumes']}条简历，{data['summary']['totalOffers']}个Offer")
        return True

    except Exception as e:
        print(f"  ❌ 处理失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("="*50)
    print("📁 Excel 监听服务启动")
    print(f"   监控目录: {WATCH_DIR}")
    print(f"   轮询间隔: {POLL_INTERVAL}秒")
    print(f"   处理脚本: {__file__}")
    print("="*50)
    print("将新的内推数据Excel文件放入监控目录，即自动处理并更新看板。")
    print("按 Ctrl+C 停止监听。")
    print()

    # 初始化：先处理已有的最新文件
    latest = get_latest_xlsx()
    if latest:
        print(f"📦 检测到已有文件: {os.path.basename(latest)}")
        process_file(latest)
        global _last_processed
        _last_processed = latest
    else:
        print("⏳ 等待新文件进入...")

    while True:
        time.sleep(POLL_INTERVAL)
        latest = get_latest_xlsx()
        if latest and latest != _last_processed:
            _last_processed = latest
            process_file(latest)
        elif latest is None:
            # 目录空了，刷新
            _last_processed = None

if __name__ == '__main__':
    main()