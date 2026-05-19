"""
腾讯校园大使内推数据看板 - Flask 后端
支持 Excel 上传 → 自动处理 → 返回看板数据
上传 Excel 后数据存于内存（云端部署时可扩展为 COS 持久化）
"""
import os
import io
import sys
import json
import time
import uuid
import traceback
import pandas as pd
from flask import Flask, request, jsonify, send_from_directory

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

app = Flask(__name__, static_folder='.')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

# ============================================================
# 配置
# ============================================================
UHR_FILE = os.environ.get(
    'UHR_FILE',
    r'C:\Users\kaiboy\Documents\UHR-高校底表.xlsx'
)

CATEGORY_MAP = {
    'S3-HR': ['AI-HR培训生', '项目实习生-人力资源', '人力资源-沟通', '人力资源培训生'],
    'CSIG-安全': ['安全技术'],
    'CSIG-销培': ['CSIG技术产品商务培训生'],
    'IEG': ['游戏发行/运营培训生', '游戏客户端开发', '游戏引擎开发', '游戏策划培训生', '项目实习生-游戏策划', '3D生成基础大模型'],
    'IEG-美术设计': ['2D角色设计', '2D场景设计', '3D场景设计', '技术美术', '3D角色设计'],
    'CDG-MA': ['投资分析'],
    'CDG-广告': ['腾讯营销管培生'],
}

OFFER_STATUSES = ['毕业生已录用', '实习已录用']
ELITE_THRESHOLD = 200
VALID_REGIONS = {'华北', '华东', '华南', '华中', '西区', '港校', '欧洲', '北美', 'OHR', 'IEG'}

# 当前看板数据（上传后写入，初始为 None 触发引导页）
CURRENT_DATA = None


# ============================================================
# 工具函数
# ============================================================
def normalize_uhr(name):
    name = str(name).strip()
    name = name.replace('（', '(').replace('）', ')')
    return name


# ============================================================
# 核心数据处理（纯函数，输入 DataFrame + UHR DataFrame）
# ============================================================
def process_data(df, df_uhr_region):
    """处理内推数据，返回看板 JSON dict"""
    # ---- 列补全 ----
    for col in ['校园大使学校', '投递项目名称', '投递岗位类', '最高学历学院', '校园大使名称', 'uhr名称']:
        if col not in df.columns:
            df[col] = ''

    print(f"  处理 {len(df)} 行，{len(df.columns)} 列")

    # ---- 校园大使名称修正映射（每次导入必须执行）----
    NAME_FIX_MAP = {
        '蔡泽厚': '虞舒璇',
        '邵圳齐': '陈炫伊',
        '陈炫伊': '卢刘泽兮',
        '卢刘泽兮': '蔡泽厚',
        '虞舒璇': '邵圳齐'
    }
    df['校园大使名称'] = df['校园大使名称'].astype(str).str.strip().replace(NAME_FIX_MAP)

    # ---- UHR / 区域 映射 ----
    school_region_map = {}
    school_uhr_map = {}
    for _, row in df_uhr_region.iterrows():
        school = str(row.get('筛选项名称', '')).strip()
        region = str(row.get('区域', '')).strip()
        uhr = str(row.get('UHR', '')).strip()
        if school and school != 'nan':
            school_region_map[school] = region
            school_uhr_map[school] = uhr

    # ---- 总览 ----
    total_resumes = len(df)
    offer_count = int(df[df['简历流程状态'].isin(OFFER_STATUSES)].shape[0])
    graduate_offer = int(df[df['简历流程状态'] == '毕业生已录用'].shape[0])
    intern_offer = int(df[df['简历流程状态'] == '实习已录用'].shape[0])

    # ---- 流程状态分布 ----
    raw = df['简历流程状态'].value_counts(dropna=False).to_dict()
    status_dist = {}
    for k, v in raw.items():
        key = '待处理' if pd.isna(k) else str(k)
        status_dist[key] = int(v)

    # ---- BG ----
    bg_counts = df['虚拟部门BG名称'].value_counts().to_dict()
    bg_data = [
        {'name': n, 'value': int(c)}
        for n, c in sorted(bg_counts.items(), key=lambda x: -x[1])
    ]

    # ---- 校园大使排名 ----
    amb_rank = df['校园大使名称'].value_counts()
    # 批量学校
    amb_school = df.drop_duplicates('校园大使名称').set_index('校园大使名称')['校园大使学校'].to_dict()
    # 批量 Offer
    offer_mask = df['简历流程状态'].isin(OFFER_STATUSES)
    amb_offers = df[offer_mask].groupby('校园大使名称').size().to_dict()

    ambassador_rank = []
    for i, (name, cnt) in enumerate(amb_rank.items(), 1):
        ambassador_rank.append({
            'rank': i,
            'name': str(name),
            'school': str(amb_school.get(name, '')),
            'resumes': int(cnt),
            'offers': int(amb_offers.get(name, 0)),
            'isElite': int(cnt) >= ELITE_THRESHOLD
        })

    elite_count = sum(1 for a in ambassador_rank if a['isElite'])

    # ---- 区域 ----
    df['区域'] = df['最高学历学校'].map(school_region_map)
    missing = df['区域'].isna()
    df.loc[missing, '区域'] = df.loc[missing, '校园大使学校'].map(school_region_map)
    df['区域'] = df['区域'].fillna('其他')
    df.loc[~df['区域'].isin(VALID_REGIONS), '区域'] = '其他'

    region_counts = df['区域'].value_counts()
    region_offers = df[offer_mask].groupby('区域').size().to_dict()
    region_rank = [
        {'rank': i, 'name': str(n), 'resumes': int(c), 'offers': int(region_offers.get(n, 0))}
        for i, (n, c) in enumerate(region_counts.items(), 1)
    ]

    # ---- UHR ----
    df['UHR_tmp'] = df['最高学历学校'].map(school_uhr_map)
    missing2 = df['UHR_tmp'].isna()
    df.loc[missing2, 'UHR_tmp'] = df.loc[missing2, '校园大使学校'].map(school_uhr_map)
    df['UHR_final'] = df['UHR_tmp'].fillna(df['uhr名称'])
    df['UHR_final'] = df['UHR_final'].apply(normalize_uhr)

    uhr_counts = df['UHR_final'].value_counts()
    uhr_offers = df[offer_mask].groupby('UHR_final').size().to_dict()
    uhr_region_map = {
        normalize_uhr(str(r['UHR'])): str(r['区域'])
        for _, r in df_uhr_region.iterrows()
    }
    uhr_rank = [
        {'rank': i, 'name': str(n), 'region': uhr_region_map.get(str(n), ''),
         'resumes': int(c), 'offers': int(uhr_offers.get(n, 0))}
        for i, (n, c) in enumerate(uhr_counts.items(), 1)
    ]

    # ---- 院校 ----
    school_counts = df['最高学历学校'].value_counts()
    school_offers = df[offer_mask].groupby('最高学历学校').size().to_dict()
    school_rank = [
        {'rank': i, 'name': str(n), 'region': school_region_map.get(str(n), '其他'),
         'resumes': int(c), 'offers': int(school_offers.get(n, 0))}
        for i, (n, c) in enumerate(school_counts.items(), 1)
        if not pd.isna(n)
    ][:100]

    # ---- 青云计划 ----
    is_qy = df['投递项目名称'].str.contains('青云', na=False)
    qy_total = int(is_qy.sum())
    qy_intern = int(df[is_qy & (df['投递项目名称'] == '青云实习')].shape[0])
    qy_grad = int(df[is_qy & (df['投递项目名称'] == '青云计划-应届生')].shape[0])
    qy_offers = int(df[is_qy & df['简历流程状态'].isin(OFFER_STATUSES)].shape[0])
    qy_graduate_offer = int(df[is_qy & (df['简历流程状态'] == '毕业生已录用')].shape[0])
    qy_intern_offer = int(df[is_qy & (df['简历流程状态'] == '实习已录用')].shape[0])
    # 青云 TOP 岗位类
    qy_df = df[is_qy]
    qy_jobs = qy_df['投递岗位类'].value_counts().head(10)
    qy_jobs_list = [{'name': str(n), 'value': int(v)} for n, v in qy_jobs.items()]

    # ---- 7个垂类 ----
    category_list = []
    for cat_name, keywords in CATEGORY_MAP.items():
        mask = pd.Series([False] * len(df))
        for kw in keywords:
            mask |= df['投递岗位类'].str.contains(kw, na=False)
        cat_resumes = int(mask.sum())
        cat_offers = int(df[mask & df['简历流程状态'].isin(OFFER_STATUSES)].shape[0])
        category_list.append({
            'name': cat_name, 'resumes': cat_resumes, 'offers': cat_offers
        })
    category_list.sort(key=lambda x: -x['resumes'])

    # ---- 组装 ----
    result = {
        'summary': {
            'totalResumes': total_resumes,
            'totalOffers': offer_count,
            'graduateOffer': graduate_offer,
            'internOffer': intern_offer,
            'totalAmbassadors': int(df['校园大使名称'].nunique()),
            'totalSchools': int(df['最高学历学校'].nunique()),
            'offerRate': round(offer_count / total_resumes * 100, 2) if total_resumes > 0 else 0,
            'eliteCount': elite_count,
            'eliteThreshold': ELITE_THRESHOLD,
            'qingyunResumes': qy_total,
            'qingyunIntern': qy_intern,
            'qingyunGrad': qy_grad,
            'qingyunOffers': qy_offers,
            'qingyunGraduateOffer': qy_graduate_offer,
            'qingyunInternOffer': qy_intern_offer,
            'qingyunOfferRate': round(qy_offers / qy_total * 100, 2) if qy_total > 0 else 0
        },
        'statusDist': status_dist,
        'bgData': bg_data,
        'categoryData': category_list,
        'ambassadorRank': ambassador_rank,
        'regionRank': region_rank,
        'uhrRank': uhr_rank,
        'schoolRank': school_rank,
        'qingyunJobs': qy_jobs_list
    }
    return result


# ============================================================
# 路由
# ============================================================
@app.route('/')
def index():
    """返回看板页面（含初始空数据引导）"""
    return send_from_directory('.', 'dashboard.html')


@app.route('/echarts.min.js')
def echarts_js():
    return send_from_directory('.', 'echarts.min.js')


@app.route('/api/upload', methods=['POST'])
def upload():
    """
    上传 Excel → 处理 → 存入内存 → 返回完整数据
    前端拿到数据后直接渲染，无需刷新页面
    """
    global CURRENT_DATA

    if 'file' not in request.files:
        return jsonify({'error': '未找到文件，请选择 .xlsx 文件'}), 400

    file = request.files['file']
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': '仅支持 .xlsx 或 .xls 文件'}), 400

    try:
        t0 = time.time()
        file_bytes = file.read()

        # ---- 找内推数据 sheet ----
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        target_sheet = None
        for name in xl.sheet_names:
            if '外部伯乐' in name or '简历推荐' in name or '伯乐' in name:
                target_sheet = name
                break
        if not target_sheet:
            # 取行数最多的 sheet
            best, best_rows = None, 0
            for name in xl.sheet_names:
                try:
                    rows = len(pd.read_excel(io.BytesIO(file_bytes), sheet_name=name))
                    if rows > best_rows:
                        best_rows = rows
                        best = name
                except Exception:
                    pass
            target_sheet = best or xl.sheet_names[0]

        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='获取名单上的这些外部伯乐的简历推荐数据_1')
        print(f"  Sheet: '{target_sheet}', 读取 {len(df)} 行")

        # ---- 读取 UHR 底表 ----
        if os.path.exists(UHR_FILE):
            with open(UHR_FILE, 'rb') as f:
                uhr_bytes = f.read()
            df_uhr = pd.read_excel(io.BytesIO(uhr_bytes), sheet_name='Sheet2')
        else:
            # 没有底表时建立空映射
            df_uhr = pd.DataFrame(columns=['筛选项名称', 'UHR', '区域'])

        # ---- 处理数据 ----
        CURRENT_DATA = process_data(df, df_uhr)

        t1 = time.time()
        print(f"  ✅ 处理完成，耗时 {t1-t0:.1f}s，{CURRENT_DATA['summary']['totalResumes']} 条简历")

        return jsonify({
            'success': True,
            'message': f"导入成功！共 {CURRENT_DATA['summary']['totalResumes']:,} 条简历，{CURRENT_DATA['summary']['totalOffers']} 个 Offer",
            'data': CURRENT_DATA,
            'summary': CURRENT_DATA['summary']
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'处理失败：{str(e)}'}), 500


@app.route('/api/data')
def get_data():
    """获取当前看板数据（初始未上传时返回引导信息）"""
    global CURRENT_DATA
    if CURRENT_DATA is None:
        return jsonify({'loaded': False, 'error': '请先上传内推数据 Excel'})
    return jsonify({'loaded': True, **CURRENT_DATA})


# ============================================================
if __name__ == '__main__':
    print("=" * 50)
    print("🚀 腾讯校园大使数据看板")
    print("   本地访问: http://localhost:8765")
    print("   上传 Excel 后即可生成看板")
    print("=" * 50)
    app.run(host='0.0.0.0', port=8765, debug=False)