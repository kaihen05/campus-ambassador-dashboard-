"""
校园大使内推数据看板 - Flask 后端
支持 Excel 上传 → 自动处理 → 返回看板数据
上传 Excel 后数据存于内存（云端部署时可扩展为 COS / S3 持久化）
"""
import os
import io
import sys
import json
import time
import uuid
import pickle
import traceback
import pandas as pd
from flask import Flask, request, jsonify, send_from_directory

# 当前看板数据持久化路径（解决 server 重启后需要重新上传的问题）
CURRENT_DATA_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'current_data.pkl')
# 周快照目录（每次上传时把"覆盖前的旧数据"按 ISO 周编号备份一份，用于周对比）
SNAPSHOT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'snapshots')
os.makedirs(SNAPSHOT_DIR, exist_ok=True)

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

app = Flask(__name__, static_folder='.')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

# ============================================================
# 路径配置（优先用项目同目录文件，云端部署友好）
# ============================================================
HERE = os.path.dirname(os.path.abspath(__file__))


def _resolve_data_file(local_name, env_key, fallback_abs):
    """优先级：同目录文件 > 环境变量 > 绝对 fallback 路径。"""
    same_dir = os.path.join(HERE, local_name)
    if os.path.exists(same_dir):
        return same_dir
    env_val = os.environ.get(env_key)
    if env_val and os.path.exists(env_val):
        return env_val
    if fallback_abs and os.path.exists(fallback_abs):
        return fallback_abs
    return same_dir  # 返回同目录路径作为占位（即使不存在）


UHR_FILE = _resolve_data_file(
    'UHR-高校底表.xlsx', 'UHR_FILE', ''
)
LAST_YEAR_FILE = _resolve_data_file('last_year_data.xlsx', 'LAST_YEAR_FILE', '')
CATEGORY_BASELINE_FILE = _resolve_data_file(
    'category_baseline.xlsx', 'CATEGORY_BASELINE_FILE', ''
)


# ============================================================
# UHR 业务覆盖配置（从外部 JSON 加载，仓库不带名单）
# ============================================================
def _load_uhr_overrides():
    """读取 uhr_overrides.json（同目录）或 UHR_OVERRIDES_FILE 环境变量指定的文件。

    JSON 结构示例（uhr_overrides.example.json）：
        {
          "region_override": {                # UHR 名 → 主辖区域
            "someuhr(姓名)": "华北"
          },
          "by_school_uhrs": {                  # 名下学校横跨海外+大陆的 UHR
            "someuhr(姓名)": {                 # 按"最高学历学校"粒度归属
              "overseas_region": "亚太",       # 海外/港澳台 → 哪个区域
              "mainland_region": "华北"        # 中国大陆     → 哪个区域
            }
          }
        }

    若文件不存在则两个字典都为空（不影响其它功能，只是不做业务硬覆盖）。
    """
    candidates = [
        os.environ.get('UHR_OVERRIDES_FILE'),
        os.path.join(HERE, 'uhr_overrides.json'),
    ]
    for path in candidates:
        if path and os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                return (
                    dict(cfg.get('region_override') or {}),
                    dict(cfg.get('by_school_uhrs') or {}),
                )
            except Exception as e:
                print(f"  ⚠️ 读取 UHR 覆盖配置失败 {path}: {e}")
    return {}, {}


UHR_REGION_OVERRIDE_CFG, UHR_BY_SCHOOL_CFG = _load_uhr_overrides()

# ============================================================
# 业务常量
# ============================================================
# 默认垂类映射（如果 category_baseline.xlsx 加载失败则使用这个）
# ⚠️ 仓库默认为空 {}，由部署方根据自家业务线在 category_map.json 或 category_baseline.xlsx 提供。
# 结构示例：
#   { "技术": ["后台开发","算法","客户端开发"],
#     "产品": ["产品经理","项目经理"], ... }
DEFAULT_CATEGORY_MAP = {}


def _load_category_map_json():
    """读 category_map.json（同目录），结构：{ 卡片名: [岗位类关键词, ...] }"""
    path = os.environ.get('CATEGORY_MAP_FILE') or os.path.join(HERE, 'category_map.json')
    if not os.path.exists(path):
        return {}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
        if isinstance(cfg, dict):
            # 简单类型校验
            return {str(k): list(v) for k, v in cfg.items() if isinstance(v, (list, tuple))}
    except Exception as e:
        print(f"  ⚠️ category_map.json 读取失败: {e}")
    return {}


# 仓库默认 DEFAULT_CATEGORY_MAP 为空，启动时再尝试用 JSON 覆盖
DEFAULT_CATEGORY_MAP.update(_load_category_map_json())

OFFER_STATUSES = ['毕业生已录用', '实习已录用']
ELITE_THRESHOLD = 200
VALID_REGIONS = {'华北', '华东', '华南', '华中', '西区', '港校', '欧洲', '北美', '亚太', 'OHR', 'IEG'}

# ============================================================
# 校名归一化 / 别名映射
#   - 解决底表与内推数据写法不一致：括号备注 / 全半角 / 大小写 / 中英文 / 空格
#   - 所有 school_region_map / school_uhr_map 查找都走 _lookup_school()
# ============================================================
import re as _re_school

# 中英文 / 简称 → 底表标准中文名（底表 Sheet2 的“筛选项名称”）
SCHOOL_ALIAS_TO_STD = {
    # 亚太
    'nanyang technological university': '南洋理工大学',
    'ntu': '南洋理工大学',
    'national university of singapore': '新加坡国立大学',
    'nus': '新加坡国立大学',
    'the university of sydney': '悉尼大学',
    'university of sydney': '悉尼大学',
    'usyd': '悉尼大学',
    'the university of melbourne': '墨尔本大学',
    'university of melbourne': '墨尔本大学',
    'unimelb': '墨尔本大学',
    'australian national university': '澳大利亚国立大学',
    'the australian national university': '澳大利亚国立大学',
    'anu': '澳大利亚国立大学',
    'university of new south wales': '新南威尔士大学',
    'the university of new south wales': '新南威尔士大学',
    'unsw': '新南威尔士大学',
    'unsw sydney': '新南威尔士大学',
}


def _normalize_school_name(s):
    """把任意写法的校名压成可比较的 key：
       去括号备注、去空格、转小写、全半角折叠
    """
    if s is None:
        return ''
    txt = str(s).strip()
    if not txt or txt.lower() in ('nan', 'none'):
        return ''
    # 去掉中英文括号及里面的备注： 墨尔本大学（澳大利亚） → 墨尔本大学
    txt = _re_school.sub(r'[（(\[][^）)\]]*[）)\]]', '', txt)
    # 全角空格 → 半角；去掉所有空白
    txt = txt.replace('\u3000', ' ')
    txt = _re_school.sub(r'\s+', '', txt)
    # 去掉常见标点
    txt = txt.replace('·', '').replace('•', '').replace('．', '.').replace('，', ',')
    return txt.lower()


def _build_school_lookup(df_uhr_region):
    """根据底表构建归一化查找表：normalized_name → (region, uhr)
    同时把 SCHOOL_ALIAS_TO_STD 的英文别名也注入。
    返回 (lookup_dict, raw_region_map_for_legacy, raw_uhr_map_for_legacy)
    """
    lookup = {}        # normalized → {'region':..., 'uhr':..., 'std': raw_name}
    raw_region = {}    # 兼容旧逻辑：原始名 → region
    raw_uhr = {}
    for _, row in df_uhr_region.iterrows():
        school = str(row.get('筛选项名称', '')).strip()
        region = str(row.get('区域', '')).strip()
        uhr = str(row.get('UHR', '')).strip()
        if not school or school.lower() == 'nan':
            continue
        raw_region[school] = region
        raw_uhr[school] = uhr
        key = _normalize_school_name(school)
        if key:
            lookup[key] = {'region': region, 'uhr': uhr, 'std': school}
    # 别名 → 标准名 → region/uhr
    for alias, std in SCHOOL_ALIAS_TO_STD.items():
        std_key = _normalize_school_name(std)
        if std_key in lookup:
            lookup[_normalize_school_name(alias)] = lookup[std_key]
    return lookup, raw_region, raw_uhr


def _lookup_school(name, school_lookup):
    """容错查找。返回 dict 或 None。"""
    if not school_lookup:
        return None
    # 1) 原值精确
    raw = '' if name is None else str(name).strip()
    if not raw:
        return None
    # 2) 归一化后查
    key = _normalize_school_name(raw)
    if key in school_lookup:
        return school_lookup[key]
    return None

# 当前看板数据（上传后写入，初始为 None 触发引导页）
CURRENT_DATA = None
# 去年同期 summary（启动时预加载）
LAST_YEAR_DATA = None
# 当前生效的垂类映射（启动时由 category_baseline.xlsx 覆盖）
CATEGORY_MAP = dict(DEFAULT_CATEGORY_MAP)


# ============================================================
# 工具函数
# ============================================================
def normalize_uhr(name):
    name = str(name).strip()
    name = name.replace('（', '(').replace('）', ')')
    return name


# ===== 港澳台/海外院校识别（用于按学校粒度归属的跨区 UHR）=====
# 判定优先级（避免 "西安大略大学(加拿大)"、"香港中文大学(深圳)" 等歧义）：
#   1) 海外 strong override（明确海外学校中文/英文关键词）→ 海外
#   2) 大陆 override（在大陆办学的港校分校等）→ 大陆
#   3) 港澳台关键词 → 海外/港澳台
#   4) 不含汉字（纯英文/拉丁）→ 海外
#   5) 否则 → 大陆
_HKMOTW_KEYWORDS = (
    '香港', '澳门', '台湾', '澳大利亚',
    'Hong Kong', 'HongKong', 'Hongkong', 'HONG KONG',
    'Macau', 'Macao', 'MACAU', 'Universidade de Macau',
    'Taiwan', 'TAIWAN',
    # 高频缩写/常见英文写法
    ' HK', '(HK', 'HKU', 'CityU', 'CUHK', 'PolyU', 'HKBU', 'HKUST',
    '岭南大学', '澳大', '港大', '港中文', '港科技', '港理工', '港浸会',
    '港城大', '港教大', '港岭南',
    '香港中文大學',
)

# === 海外院校 strong override（中文/英文常见海外大学关键字，命中即海外）===
# 谨慎：仅放进 "中国大陆没有同名/相似学校" 的关键字。
# 例：'西北大学' 中国有西北大学(西安)，英文 Northwestern 才是美国 → 只用 'Northwestern' 而非 '西北大学'
_OVERSEAS_STRONG_KEYWORDS = (
    # === 港澳台扩充（中文里没的写法）===
    'Chinese University of Hong Kong', 'University of Hong Kong',
    'Hong Kong Polytechnic', 'Hong Kong Baptist', 'Lingnan University',
    'Hong Kong University of Science', 'Education University of Hong Kong',
    # === 新加坡 ===
    '新加坡', 'Singapore', '南洋理工', 'Nanyang Technological', 'NTU',
    'National University of Singapore', 'NUS', 'Singapore Management',
    'SMU', 'Singapore University of Technology',
    # === 澳洲 / 新西兰 ===
    '悉尼', '墨尔本', '昆士兰', '新南威尔士', '阿德莱德', '莫纳什', '蒙纳士',
    '麦考瑞', '塔斯马尼亚', '伍伦贡', '迪肯', '科廷', '皇家墨尔本',
    'Sydney', 'Melbourne', 'Queensland', 'UNSW', 'New South Wales',
    'Adelaide', 'Monash', 'Macquarie', 'Tasmania', 'Wollongong',
    'Australian National University', 'ANU', 'University of Western Australia',
    'UWA', 'Auckland', '奥克兰', 'Otago', '惠灵顿', 'Victoria University of Wellington',
    # === 英国 ===
    '伦敦', '剑桥', '牛津', '爱丁堡', '曼彻斯特', '帝国理工', '华威', '利兹',
    '格拉斯哥', '伯明翰', '南安普顿', '布里斯托', '杜伦', '兰卡斯特', '诺丁汉',
    '谢菲尔德', '埃克塞特', '利物浦', '约克大学', '巴斯', '雷丁', '萨里',
    '苏塞克斯', '东英吉利', '纽卡斯尔', '卡迪夫', '贝尔法斯特', '皇后大学',
    '伦敦国王', '伦敦大学学院', '伦敦政治经济',
    'London', 'Cambridge', 'Oxford', 'Edinburgh', 'Manchester', 'Imperial College',
    'Warwick', 'Leeds', 'Glasgow', 'Birmingham', 'Southampton', 'Bristol',
    'Durham', 'Lancaster', 'Sheffield', 'Exeter', 'Liverpool', 'York',
    'Bath', 'Reading', 'Surrey', 'Sussex', 'Newcastle', 'Cardiff', 'UCL', 'KCL',
    'King\'s College', 'LSE', 'School of Economics', 'St Andrews', '圣安德鲁斯',
    'Queen Mary', '玛丽女王', 'Royal Holloway', 'Goldsmiths', 'SOAS',
    # === 欧洲（除英国）===
    '慕尼黑', '柏林', '海德堡', '亚琛', '法兰克福', '波恩', '科隆',
    '巴黎', '索邦', '里昂', '马赛', '图卢兹',
    '苏黎世', '日内瓦', '洛桑', '巴塞尔', '伯尔尼',
    '阿姆斯特丹', '代尔夫特', '埃因霍温', '鹿特丹', '莱顿',
    '哥本哈根', '隆德', '斯德哥尔摩', '奥斯陆', '赫尔辛基',
    '都柏林', '科克大学', '马德里', '巴塞罗那', '罗马', '米兰', '博洛尼亚',
    '布鲁塞尔', '鲁汶', '维也纳', '布拉格', '华沙',
    'Munich', 'Berlin', 'Heidelberg', 'Aachen', 'Frankfurt',
    'Sorbonne', 'Sciences Po', 'HEC Paris', 'INSEAD',
    'ETH Zurich', 'EPFL', 'Lausanne', 'Geneva',
    'Amsterdam', 'Delft', 'Eindhoven', 'Erasmus', 'Leiden', 'Utrecht',
    'Copenhagen', 'Lund', 'Stockholm', 'Karolinska', 'Uppsala',
    'Trinity College Dublin', 'University College Dublin',
    # === 北美（美国）===
    '哈佛', '麻省理工', '斯坦福', '伯克利', '普林斯顿', '耶鲁',
    '哥伦比亚大学', '哥大', '康奈尔', '宾夕法尼亚大学', '宾大',
    '卡内基梅隆', '卡耐基梅隆', '卡耐基', '芝加哥大学', '约翰霍普金斯',
    '杜克大学', '杜克', '加州理工', '西北大学(美国)', 'Northwestern',
    '密歇根大学', '密西根大学', '威斯康星麦迪逊', '伊利诺伊', '德州奥斯汀',
    '德克萨斯', '北卡', '佐治亚理工', '佐治亚大学', '弗吉尼亚大学',
    '马里兰大学', '宾州州立', '俄亥俄', '波士顿大学', '波士顿学院',
    '东北大学(美国)', '塔夫茨', '布朗大学', '达特茅斯', '纽约大学',
    '南加州大学', '加州大学', '加州州立', '加州大学洛杉矶', '加州大学圣地亚哥',
    '加州大学戴维斯', '加州大学欧文', '加州大学圣芭芭拉', '加州大学河滨',
    '罗格斯', '雪城大学', '罗切斯特', '范德比尔特', '埃默里', '莱斯大学',
    '圣路易斯华盛顿', '华盛顿大学(美国)', '亚利桑那州立', '亚利桑那大学',
    '匹兹堡大学', '迈阿密大学', '佛罗里达大学', '德雷塞尔',
    'Harvard', 'MIT', 'Massachusetts Institute', 'Stanford', 'Berkeley',
    'Princeton', 'Yale', 'Columbia University', 'Cornell', 'UPenn',
    'Pennsylvania', 'Carnegie Mellon', 'CMU', 'Chicago', 'Johns Hopkins', 'JHU',
    'Duke', 'Caltech', 'California Institute of Technology',
    'University of Michigan', 'Michigan-Ann Arbor', 'Wisconsin-Madison',
    'Illinois Urbana', 'UIUC', 'Texas at Austin', 'UT Austin',
    'North Carolina', 'UNC', 'Georgia Tech', 'Virginia',
    'Maryland', 'Penn State', 'Ohio State', 'Boston University',
    'Boston College', 'Northeastern University', 'Tufts', 'Brown University',
    'Dartmouth', 'NYU', 'New York University',
    'USC', 'University of Southern California',
    'UCLA', 'UCSD', 'UC San Diego', 'UC Davis', 'UC Irvine', 'UC Santa Barbara',
    'Rutgers', 'Syracuse', 'Rochester', 'Vanderbilt', 'Emory', 'Rice University',
    'Washington University in St', 'Washington University-St', 'WashU',
    'Arizona State', 'University of Arizona', 'Pittsburgh', 'Miami',
    'University of Florida', 'Florida State', 'Drexel',
    # === 北美（加拿大）===
    '多伦多大学', '麦吉尔', '不列颠哥伦比亚', '英属哥伦比亚', '滑铁卢',
    '麦克马斯特', '西安大略', '韦仕敦', '皇后大学(加拿大)', '阿尔伯塔',
    '渥太华', '西蒙菲莎', '西蒙弗雷泽', '维多利亚大学(加拿大)', '康考迪亚',
    'Toronto', 'McGill', 'British Columbia', 'UBC', 'Waterloo',
    'McMaster', 'Western Ontario', 'Queen\'s University', 'Alberta',
    'Ottawa', 'Simon Fraser', 'SFU', 'Concordia',
    # === 日韩 ===
    '东京大学', '京都大学', '大阪大学', '早稻田', '庆应', '一桥', '东工大',
    '北海道大学', '名古屋大学', '九州大学', '东北大学(日本)',
    '首尔大学', '高丽大学', '延世大学', '成均馆', 'KAIST', '浦项',
    'Tokyo', 'Kyoto', 'Osaka', 'Waseda', 'Keio', 'Hitotsubashi',
    'Hokkaido', 'Nagoya', 'Kyushu', 'Tohoku',
    'Seoul National', 'Korea University', 'Yonsei', 'Sungkyunkwan',
    # === 东南亚/南亚/中东 ===
    '马来亚', '马来西亚', '泰国', '清迈', '朱拉隆功', '玛希隆',
    '印度理工', '印度商学院', '印度尼西亚', '菲律宾', '越南',
    '阿拉伯', '阿布扎比', '迪拜', '以色列', '希伯来大学', '特拉维夫',
    'Malaya', 'Malaysia', 'Chulalongkorn', 'Mahidol', 'Indonesia',
    'IIT', 'Indian Institute', 'Hebrew University', 'Tel Aviv',
)

# 一些容易误判的"含英文但其实是中国大陆校"白名单（这些走"华北"）
# 注意：放在 _OVERSEAS_STRONG_KEYWORDS 之后判断，避免 "西安大略" 被 '西安' 误拦
_MAINLAND_OVERRIDE_KEYWORDS = (
    '中国', '清华', '北京', '上海', '深圳', '广州', '成都', '武汉', '西安', '南京', '杭州',
    '宁波诺丁汉', '昆山杜克', '西交利物浦', '香港中文大学(深圳)', '香港中文大学（深圳）',
    '北师香港浸会', '北京师范大学-香港浸会', '北京师范大学 - 香港浸会',
    '香港科技大学（广州）', '香港科技大学(广州)',
    '深圳北理莫斯科',
)


def _is_overseas_or_hkmotw_school(name):
    """判断学校是否属于"港澳台/海外"。"""
    if not name:
        return False
    s = str(name).strip()
    if not s or s.lower() in ('nan', 'none', '<na>'):
        return False
    s_lower = s.lower()
    # 0) 中外合办校（在大陆办学）特殊白名单 → 大陆
    #    必须先于 strong override 检查，避免 "上海纽约大学" 被 'NYU/纽约大学' 误判海外
    _SINO_FOREIGN_MAINLAND = (
        '上海纽约大学', '昆山杜克大学', '宁波诺丁汉', '西交利物浦',
        '香港中文大学(深圳)', '香港中文大学（深圳）',
        '香港科技大学（广州）', '香港科技大学(广州)',
        '北师香港浸会', '北京师范大学-香港浸会', '北京师范大学 - 香港浸会',
        '深圳北理莫斯科', '广东以色列', '温州肯恩',
        '上海纽约', 'NYU Shanghai', 'Duke Kunshan', 'Nottingham Ningbo',
        'Xi\'an Jiaotong-Liverpool', 'Chinese University of Hong Kong, Shenzhen',
    )
    for kw in _SINO_FOREIGN_MAINLAND:
        if kw and (kw in s or kw.lower() in s_lower):
            return False
    # 1) 海外 strong override（先于大陆 override，避免"西安大略"被'西安'误拦）
    for kw in _OVERSEAS_STRONG_KEYWORDS:
        if kw and kw.lower() in s_lower:
            return True
    # 2) 大陆 override（如"香港中文大学（深圳）"）
    for kw in _MAINLAND_OVERRIDE_KEYWORDS:
        if kw in s:
            return False
    # 3) 命中港澳台关键词 → 海外/港澳台
    for kw in _HKMOTW_KEYWORDS:
        if kw in s:
            return True
    # 4) 不含任何汉字（纯英文/拉丁字符）→ 视为海外
    has_chinese = any('\u4e00' <= ch <= '\u9fff' for ch in s)
    if not has_chinese:
        return True
    # 5) 含汉字且未命中任何海外关键字 → 中国大陆
    return False


def load_uhr_df():
    """读取 UHR 底表，返回 DataFrame。文件缺失时返回空表。"""
    if os.path.exists(UHR_FILE):
        try:
            return pd.read_excel(UHR_FILE, sheet_name='Sheet2')
        except Exception as e:
            print(f"  ⚠️ UHR 底表读取失败: {e}")
    return pd.DataFrame(columns=['筛选项名称', 'UHR', '区域'])


# baseline xlsx 里的「部门」 → 当前看板的垂类卡片名
# ⚠️ 仓库默认为空 {}，由部署方根据自家业务线在 category_dept_map.json 提供。
# 结构示例：
#   { "技术线": "技术",
#     "产品线": "产品",
#     ... }
BASELINE_DEPT_TO_CATEGORY = {}

_dept_map_path = os.environ.get('CATEGORY_DEPT_MAP_FILE') or os.path.join(HERE, 'category_dept_map.json')
if os.path.exists(_dept_map_path):
    try:
        with open(_dept_map_path, 'r', encoding='utf-8') as _f:
            _cfg = json.load(_f)
        if isinstance(_cfg, dict):
            BASELINE_DEPT_TO_CATEGORY.update({str(k): str(v) for k, v in _cfg.items()})
    except Exception as _e:
        print(f"  ⚠️ category_dept_map.json 读取失败: {_e}")


def load_category_map():
    """
    从 category_baseline.xlsx（表头在第 2 行）扩充 CATEGORY_MAP 的关键词。
    保留默认垂类卡片名不变，只把 baseline 里更细分的 "投递岗位类" 关键词加进去。
    """
    global CATEGORY_MAP
    if not os.path.exists(CATEGORY_BASELINE_FILE):
        print(f"  ℹ️ 未找到 category_baseline.xlsx，使用默认 CATEGORY_MAP")
        return
    try:
        df = pd.read_excel(CATEGORY_BASELINE_FILE, header=1)
        if not {'部门', '投递岗位类'}.issubset(df.columns):
            print(f"  ⚠️ category_baseline 列不符（需 部门/投递岗位类），保留默认；现有列={list(df.columns)}")
            return
        added = 0
        for _, row in df.iterrows():
            dept_raw = str(row.get('部门', '')).strip()
            job = str(row.get('投递岗位类', '')).strip()
            if not dept_raw or dept_raw == 'nan' or not job or job == 'nan':
                continue
            cat = BASELINE_DEPT_TO_CATEGORY.get(dept_raw)
            if not cat:
                continue
            CATEGORY_MAP.setdefault(cat, [])
            if job not in CATEGORY_MAP[cat]:
                CATEGORY_MAP[cat].append(job)
                added += 1
        print(f"  ✅ 已从 category_baseline.xlsx 扩充 {added} 个岗位类关键词")
    except Exception as e:
        print(f"  ⚠️ category_baseline 加载失败: {e}")


# 去年数据 → 今年列名 的映射
# 注：去年「底表（全）」里 UHR 字段名为「外部伯乐对接uhr」，必须映射到今年规范的 'uhr名称'
# 否则 process_data 里 UHR_tmp 全靠学校反查，匹配不上的就会塞成空字符串，
# 在 uhrRank 里堆成一条 name='' 的幽灵记录（曾经表现为：By UHR—去年 第3名 简历4193/Offer67 但没名字）
LAST_YEAR_COL_MAP = {
    '外部伯乐姓名': '校园大使名称',
    '外部伯乐学校': '校园大使学校',
    '应聘项目': '投递项目名称',
    '职位小类': '投递岗位类',
    '简历状态': '简历流程状态',
    '外部伯乐对接uhr': 'uhr名称',
}


def load_last_year_data():
    """启动时预处理去年数据，返回与 process_data 同结构的 dict。带 pickle 缓存。"""
    global LAST_YEAR_DATA
    if not os.path.exists(LAST_YEAR_FILE):
        print(f"  ℹ️ 未找到 last_year_data.xlsx，跳过同比")
        return

    # 缓存：last_year_data.pkl，依赖 xlsx mtime 失效
    cache_path = os.path.join(HERE, 'last_year_data.pkl')
    try:
        if os.path.exists(cache_path) and os.path.getmtime(cache_path) >= os.path.getmtime(LAST_YEAR_FILE):
            import pickle
            with open(cache_path, 'rb') as f:
                LAST_YEAR_DATA = pickle.load(f)
            print(f"  ⚡ 命中去年数据缓存：简历 {LAST_YEAR_DATA['summary']['totalResumes']} / Offer {LAST_YEAR_DATA['summary']['totalOffers']}")
            return
    except Exception as e:
        print(f"  ⚠️ 缓存读取失败，重算: {e}")

    try:
        t0 = time.time()
        # 优先读 “底表（全）”，回退到行数最多的 sheet
        xl = pd.ExcelFile(LAST_YEAR_FILE)
        target = None
        for name in xl.sheet_names:
            if '底表（全）' in name or '底表(全)' in name:
                target = name
                break
        if not target:
            best, best_rows = None, 0
            for name in xl.sheet_names:
                try:
                    rows = len(pd.read_excel(LAST_YEAR_FILE, sheet_name=name))
                    if rows > best_rows:
                        best_rows, best = rows, name
                except Exception:
                    pass
            target = best or xl.sheet_names[0]

        df = pd.read_excel(LAST_YEAR_FILE, sheet_name=target)
        print(f"  📚 去年数据 sheet='{target}'，{len(df)} 行")

        # 列名映射 → 今年规范
        df = df.rename(columns=LAST_YEAR_COL_MAP)

        # 补齐今年代码必需的列
        for col in ['校园大使名称', '校园大使学校', '最高学历学校', '最高学历学院',
                    '投递项目名称', '投递岗位类', '简历流程状态', '虚拟部门BG名称', 'uhr名称']:
            if col not in df.columns:
                df[col] = ''

        # 学校缺失时回填
        df['最高学历学校'] = df['最高学历学校'].fillna('').replace('', pd.NA)
        df['最高学历学校'] = df['最高学历学校'].fillna(df['校园大使学校'])
        df['最高学历学校'] = df['最高学历学校'].fillna('').astype(str)

        # === 关键：用「洗数据（排名）」sheet 的"大使→UHR"权威映射覆盖 uhr名称 ===
        # 用户要求严格按这张表的 UHR 列计算去年 UHR 排名（254 名大使 → 13 个 UHR）
        # 「底表（全）」里的 外部伯乐对接uhr 字段并不是按"大使→UHR"最终归属算的，会出现同一大使
        # 挂在多个 UHR 名下的情况，导致 UHR 排名失真。
        try:
            rank_df = pd.read_excel(LAST_YEAR_FILE, sheet_name='洗数据（排名）', header=0)
            rank_df.columns = ['大使姓名', '简历数', 'UHR']
            rank_df = rank_df.iloc[1:].reset_index(drop=True)  # 去掉首行汇总
            rank_df['大使姓名'] = rank_df['大使姓名'].astype(str).str.strip()
            rank_df['UHR'] = rank_df['UHR'].astype(str).str.strip()
            # 去重保险（理论上唯一）
            rank_df = rank_df.drop_duplicates(subset='大使姓名', keep='first')
            ambassador_to_uhr = dict(zip(rank_df['大使姓名'], rank_df['UHR']))
            print(f"  📑 「洗数据（排名）」加载成功：{len(ambassador_to_uhr)} 名大使 → {rank_df['UHR'].nunique()} 个 UHR")
            # 用大使姓名映射 UHR，覆盖原值；找不到的保留原 uhr名称（极少数）
            mapped = df['校园大使名称'].astype(str).str.strip().map(ambassador_to_uhr)
            hit = mapped.notna() & (mapped.astype(str).str.strip() != '')
            df.loc[hit, 'uhr名称'] = mapped[hit].values
            print(f"  ✅ 大使→UHR 命中 {int(hit.sum())} / {len(df)} 条简历")
            # 标记：去年数据的 uhr名称 已经是权威值，process_data 应直接信任，跳过"学校→UHR"反查
            df.attrs['uhr_is_authoritative'] = True
        except Exception as e:
            print(f"  ⚠️ 「洗数据（排名）」加载失败，回退到学校反查: {e}")

        # === 关键：用「洗数据（offer）」+「底表（offer）」 严格按 sheet 重写 offer 状态 ===
        # 用户要求："严格按照 Sheet 洗数据（offer）中的 UHR 列进行计算"
        # 「底表（全）」原本用「简历状态」=='实习已录用'/'毕业生已录用' 判断 offer，但权威总数（1927）
        # 包含 OFFER报批中 / OFFER待报批 / OFFER报批流程放弃 等状态，这些状态在 OFFER_STATUSES 里命中不到。
        # 解决方案：用「底表（offer）」sheet 的 1927 个 简历id 作为去年 offer 的权威集合，
        # 给 df 加一列 _is_offer_authoritative=True 用于权威判断；
        # 同时用「对应关系（替换后）」sheet 把 外部伯乐id 映射到标准化大使名（清洗别名/化名），
        # 确保按 大使姓名 聚合时与「洗数据（offer）」完全一致。
        try:
            off_df = pd.read_excel(LAST_YEAR_FILE, sheet_name='底表（offer）')
            off_id_set = set(off_df['简历id'].dropna().astype(str).tolist())
            # 简历id 列：底表（全）也叫「简历id」（rename 后未变）
            id_col = '简历id' if '简历id' in df.columns else None
            if id_col:
                df['_is_offer_authoritative'] = df[id_col].astype(str).isin(off_id_set)
                print(f"  📑 「底表（offer）」加载成功：{len(off_id_set)} 个权威 offer 简历id；命中 {int(df['_is_offer_authoritative'].sum())} / {len(df)} 行")
            else:
                print(f"  ⚠️ 去年数据无「简历id」列，无法应用权威 offer 标记")
                df['_is_offer_authoritative'] = False

            # 用 对应关系（替换后） 的 bid → name 做 大使姓名标准化
            # 仅对「offer 简历」做姓名标准化 + UHR 重映射，非 offer 简历保持原状
            # 这样能保证：① UHR offer 数 严格匹配「洗数据（offer）」  ② UHR 简历数 严格匹配「洗数据（排名）」
            try:
                corr = pd.read_excel(LAST_YEAR_FILE, sheet_name='对应关系（替换后）', header=0)
                bid_to_name = {}
                for _, row in corr.iterrows():
                    bid = row.get('bid')
                    nm = row.get('name')
                    if pd.isna(bid) or pd.isna(nm):
                        continue
                    bid_to_name[float(bid)] = str(nm).strip()
                if '外部伯乐id' in df.columns and bid_to_name and '_is_offer_authoritative' in df.columns:
                    bid_series = pd.to_numeric(df['外部伯乐id'], errors='coerce')
                    std_name = bid_series.map(bid_to_name)
                    # 仅在 offer 简历上覆盖大使姓名（非 offer 简历保持原状）
                    offer_rows = df['_is_offer_authoritative'].astype(bool)
                    overwrite_mask = offer_rows & std_name.notna()
                    df.loc[overwrite_mask, '校园大使名称'] = std_name[overwrite_mask].values
                    print(f"  📑 「对应关系（替换后）」加载成功：{len(bid_to_name)} 条 bid→标准名；"
                          f"仅对 offer 简历重写 {int(overwrite_mask.sum())} / {int(offer_rows.sum())} 行")
                    # 标准化后必须重新映射 offer 简历的 UHR（因为大使名变了）
                    try:
                        mapped2 = df.loc[overwrite_mask, '校园大使名称'].astype(str).str.strip().map(ambassador_to_uhr)
                        hit2 = mapped2.notna() & (mapped2.astype(str).str.strip() != '')
                        # 把命中的索引取出来，写回 uhr名称
                        hit_idx = mapped2[hit2].index
                        df.loc[hit_idx, 'uhr名称'] = mapped2.loc[hit_idx].values
                        print(f"  ✅ 标准化后 offer 简历 大使→UHR 重映射命中 {int(hit2.sum())} 条")
                    except Exception:
                        pass
            except Exception as e:
                print(f"  ⚠️ 「对应关系（替换后）」加载失败: {e}")

            # 标记：去年数据的 offer 已用权威集合标记，process_data 应优先用 _is_offer_authoritative
            df.attrs['offer_is_authoritative'] = True
        except Exception as e:
            print(f"  ⚠️ 「底表（offer）」加载失败，回退到 简历状态 判断: {e}")
            df['_is_offer_authoritative'] = False

        df_uhr = load_uhr_df()
        LAST_YEAR_DATA = process_data(df, df_uhr)
        # 写缓存
        try:
            import pickle
            with open(cache_path, 'wb') as f:
                pickle.dump(LAST_YEAR_DATA, f)
        except Exception as e:
            print(f"  ⚠️ 缓存写入失败: {e}")
        print(f"  ✅ 去年数据预处理完成，耗时 {time.time()-t0:.1f}s，"
              f"简历 {LAST_YEAR_DATA['summary']['totalResumes']} / Offer {LAST_YEAR_DATA['summary']['totalOffers']}")
    except Exception as e:
        traceback.print_exc()
        print(f"  ⚠️ 去年数据加载失败: {e}")


# ============================================================
# 核心数据处理（纯函数，输入 DataFrame + UHR DataFrame）
# ============================================================
def process_data(df, df_uhr_region):
    """处理内推数据，返回看板 JSON dict"""
    # ---- 列补全 ----
    for col in ['校园大使学校', '投递项目名称', '投递岗位类', '最高学历学院', '校园大使名称', 'uhr名称', '虚拟部门BG名称', '最高学历学校', '简历流程状态']:
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
    # 构建鲁棒查找表：归一化名 → {region, uhr, std}
    # 同时保留旧的 raw map，仅用于兼容（实际查找全部走 _lookup_school）
    school_lookup, school_region_map, school_uhr_map = _build_school_lookup(df_uhr_region)

    def _region_of(name):
        hit = _lookup_school(name, school_lookup)
        return hit['region'] if hit else None

    def _uhr_of(name):
        hit = _lookup_school(name, school_lookup)
        return hit['uhr'] if hit else None

    # ---- 总览 ----
    total_resumes = len(df)
    # 权威 offer 标记：去年数据通过「底表（offer）」sheet 标记了 _is_offer_authoritative
    # 此时所有 offer 判定一律以该列为准；今年/无标记数据继续用 OFFER_STATUSES.isin 判断
    if '_is_offer_authoritative' in df.columns and df.attrs.get('offer_is_authoritative'):
        _offer_bool = df['_is_offer_authoritative'].astype(bool)
        print(f"  ✅ 使用权威 offer 标记（来自「底表（offer）」），共 {int(_offer_bool.sum())} 个 offer")
    else:
        _offer_bool = df['简历流程状态'].isin(OFFER_STATUSES)
    offer_count = int(_offer_bool.sum())
    graduate_offer = int((_offer_bool & (df['简历流程状态'] == '毕业生已录用')).sum())
    intern_offer = int((_offer_bool & (df['简历流程状态'] == '实习已录用')).sum())

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
    offer_mask = _offer_bool
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
    # 去年数据：UHR 是权威字段（来自「洗数据（排名）」），区域应按 UHR→区域 归属，
    # 而不是按候选人学校（去年表「最高学历学校」字段大量缺失/不规范，会把同一区域大 UHR
    # 的几千条简历错误打散到全国各区，主辖区严重低估）。
    # 今年数据：UHR 字段不可信，仍按候选人学校→区域 映射（原逻辑）。
    if df.attrs.get('uhr_is_authoritative'):
        # 业务侧覆盖（从外部 uhr_overrides.json 加载，仓库不带名单）：
        # 1) region_override：当一个 UHR 在 Sheet2 里跨多区域时，按"主辖区"硬归属。
        # 2) by_school_uhrs：名下学校横跨海外+大陆的 UHR，按"最高学历学校粒度"区分
        #    海外院校（含港澳台）→ overseas_region；中国大陆 → mainland_region。
        UHR_REGION_OVERRIDE = dict(UHR_REGION_OVERRIDE_CFG)
        # 用 UHR→区域 字典（normalize_uhr 已统一全/半角括号）
        # 当一个 UHR 在 Sheet2 里出现多个区域时，按"学校数最多"归属（避免 dict
        # 推导式"最后写入胜出"导致跨区 UHR 的简历被错误归并）。
        from collections import Counter
        uhr_region_counter = {}  # uhr_norm -> Counter({region: 学校数})
        for _, r in df_uhr_region.iterrows():
            u = normalize_uhr(str(r.get('UHR', '')))
            reg = str(r.get('区域', '')).strip()
            if not u or not reg:
                continue
            uhr_region_counter.setdefault(u, Counter())[reg] += 1
        uhr_to_region = {
            u: cnt.most_common(1)[0][0] for u, cnt in uhr_region_counter.items()
        }
        # 应用业务侧 override
        for uhr_raw, region in UHR_REGION_OVERRIDE.items():
            uhr_to_region[normalize_uhr(uhr_raw)] = region
        print(f"  📋 [去年] UHR→区域 字典构建完成（共 {len(uhr_to_region)} 个 UHR），override 数：{len(UHR_REGION_OVERRIDE)}")
        df['区域'] = df['uhr名称'].apply(normalize_uhr).map(uhr_to_region)

        # ===== 名下学校横跨海外+大陆的 UHR：按"最高学历学校"粒度判定 =====
        for uhr_raw, cfg in UHR_BY_SCHOOL_CFG.items():
            ovs_region = cfg.get('overseas_region')
            mlb_region = cfg.get('mainland_region')
            if not (ovs_region and mlb_region):
                continue
            u_norm = normalize_uhr(uhr_raw)
            mask = df['uhr名称'].apply(normalize_uhr) == u_norm
            if not mask.any():
                continue
            schools = df.loc[mask, '最高学历学校'].fillna('').astype(str)
            is_oversea = schools.apply(_is_overseas_or_hkmotw_school)
            df.loc[mask & is_oversea, '区域'] = ovs_region
            df.loc[mask & ~is_oversea, '区域'] = mlb_region
            n_total = int(mask.sum())
            n_ovs = int((mask & is_oversea).sum())
            n_mlb = int((mask & ~is_oversea).sum())
            print(f"  🎯 [去年] UHR(by-school) 共 {n_total} 条 → {ovs_region} {n_ovs} 条 / {mlb_region} {n_mlb} 条")

        # UHR 缺失或无区域映射的简历，回退按学校映射；再不行打"其他"
        missing = df['区域'].isna() | (df['区域'].astype(str).str.strip() == '')
        df.loc[missing, '区域'] = df.loc[missing, '最高学历学校'].map(_region_of)
        still_missing = df['区域'].isna()
        df.loc[still_missing, '区域'] = df.loc[still_missing, '校园大使学校'].map(_region_of)
        df['区域'] = df['区域'].fillna('其他')
        print(f"  ✅ [去年] 区域按 UHR 权威映射，共命中 {(~missing).sum()} / {len(df)} 条简历")
    else:
        df['区域'] = df['最高学历学校'].map(_region_of)
        missing = df['区域'].isna()
        df.loc[missing, '区域'] = df.loc[missing, '校园大使学校'].map(_region_of)
        df['区域'] = df['区域'].fillna('其他')
    # 标记原因：在白名单外的（含空串、其他、海外、未分类等）统一归为"其他"
    raw_region_before = df['区域'].copy()
    df.loc[~df['区域'].isin(VALID_REGIONS), '区域'] = '其他'

    # ---- 诊断：哪些学校落进了"其他" ----
    other_mask = df['区域'] == '其他'
    other_df = df[other_mask].copy()
    other_df['_raw_region'] = raw_region_before[other_mask]
    # 学校->(命中条数, 原始区域值集合, 学校字段来源)
    other_breakdown = []
    if len(other_df) > 0:
        # 优先用最高学历学校，否则用校园大使学校
        other_df['_school_used'] = other_df['最高学历学校'].fillna('').astype(str).str.strip()
        empty_mask = other_df['_school_used'].isin(['', 'nan', 'NaN'])
        other_df.loc[empty_mask, '_school_used'] = (
            other_df.loc[empty_mask, '校园大使学校'].fillna('').astype(str).str.strip()
        )
        other_df.loc[other_df['_school_used'].isin(['', 'nan', 'NaN']), '_school_used'] = '(空学校字段)'

        grouped = other_df.groupby('_school_used').agg(
            count=('_school_used', 'size'),
            raw_regions=('_raw_region', lambda s: sorted(set(str(x) for x in s if str(x) not in ('nan', 'NaN'))))
        ).reset_index().sort_values('count', ascending=False)
        for _, row in grouped.iterrows():
            raws = row['raw_regions']
            if not raws:
                reason = '映射表中无该学校（未匹配）'
            elif raws == ['']:
                reason = '映射表中区域字段为空'
            else:
                reason = f'映射区域不在白名单：{raws}'
            other_breakdown.append({
                'school': row['_school_used'],
                'count': int(row['count']),
                'reason': reason
            })

        # 打印到控制台 / 日志
        print(f"\n========== [区域诊断] 共 {len(other_df)} 条简历落入「其他」，涉及 {len(other_breakdown)} 所学校 ==========")
        for item in other_breakdown[:50]:
            print(f"  · {item['school']:<30} {item['count']:>5} 条   {item['reason']}")
        if len(other_breakdown) > 50:
            print(f"  ... 其余 {len(other_breakdown) - 50} 所学校见 /api/region_other 接口")
        print("=" * 80 + "\n")

    # ---- 诊断：亚太命中情况 ----
    ap_mask = df['区域'] == '亚太'
    if ap_mask.any():
        ap_schools = df.loc[ap_mask, '最高学历学校'].fillna('').astype(str).value_counts()
        print(f"========== [亚太命中] {int(ap_mask.sum())} 条简历，{len(ap_schools)} 所学校 ==========")
        for n, c in ap_schools.items():
            print(f"  ✓ {n:<35} {int(c):>5} 条")
        print("=" * 80 + "\n")
    else:
        print("========== [亚太命中] ⚠️ 0 条！请检查 UHR-高校底表 Sheet2 中 7 所亚太学校的写法是否与内推数据一致 ==========\n")

    region_counts = df['区域'].value_counts()
    region_offers = df[offer_mask].groupby('区域').size().to_dict()
    region_rank = [
        {'rank': i, 'name': str(n), 'resumes': int(c), 'offers': int(region_offers.get(n, 0))}
        for i, (n, c) in enumerate(region_counts.items(), 1)
    ]

    # ---- UHR ----
    # 1) 先用底表（学校 → UHR）查找；2) 查不到回退到原始 'uhr名称'（去年数据来自「外部伯乐对接uhr」）
    # 注：_uhr_of 在底表里 UHR 为空时会返回 ''；fillna 不会替换 ''，所以先把空串/'nan' 全部转成 NaN
    def _blank_to_nan(x):
        if x is None:
            return pd.NA
        s = str(x).strip()
        if s == '' or s.lower() in ('nan', 'none', '<na>'):
            return pd.NA
        return s

    # 去年数据已通过「洗数据（排名）」sheet 把 uhr名称 字段重写为权威值
    # 此时跳过"学校→UHR"反查，直接信任 uhr名称（这是用户的明确要求）
    if df.attrs.get('uhr_is_authoritative'):
        df['UHR_final'] = df['uhr名称'].map(_blank_to_nan)
    else:
        df['UHR_tmp'] = df['最高学历学校'].map(_uhr_of).map(_blank_to_nan)
        missing2 = df['UHR_tmp'].isna()
        df.loc[missing2, 'UHR_tmp'] = df.loc[missing2, '校园大使学校'].map(_uhr_of).map(_blank_to_nan)
        df['UHR_final'] = df['UHR_tmp'].fillna(df['uhr名称'].map(_blank_to_nan))
    df['UHR_final'] = df['UHR_final'].apply(normalize_uhr)
    # 仍为空的（极少数）丢给一个统一标签，避免污染排行
    df.loc[df['UHR_final'].isin(['', 'nan', 'None', '<NA>']) | df['UHR_final'].isna(), 'UHR_final'] = '未知UHR'

    uhr_counts = df['UHR_final'].value_counts()
    # uhrRank 输出时剔除占位标签，前端不展示
    uhr_counts = uhr_counts[~uhr_counts.index.isin(['未知UHR'])]
    uhr_offers = df[offer_mask].groupby('UHR_final').size().to_dict()
    # uhr_region_map: 跟前面"区域映射"逻辑严格一致（按学校数最多 + 业务 override），
    # 否则 uhrRank 卡片里跨区 UHR 又会显示错误的主辖区。
    from collections import Counter as _Counter
    _uhr_region_counter = {}
    for _, _r in df_uhr_region.iterrows():
        _u = normalize_uhr(str(_r.get('UHR', '')))
        _reg = str(_r.get('区域', '')).strip()
        if _u and _reg:
            _uhr_region_counter.setdefault(_u, _Counter())[_reg] += 1
    uhr_region_map = {
        _u: _cnt.most_common(1)[0][0] for _u, _cnt in _uhr_region_counter.items()
    }
    # 业务侧覆盖（与上方 process_data 区域映射保持一致）：
    # 1) UHR_REGION_OVERRIDE_CFG：直接固定区域；
    # 2) UHR_BY_SCHOOL_CFG：跨海外+大陆的 UHR，卡片标签默认展示其 overseas_region
    #    （因为这类 UHR 一般核心管辖在海外/港澳台），明细按学校粒度展示由 process_data 处理。
    _UHR_REGION_OVERRIDE = dict(UHR_REGION_OVERRIDE_CFG)
    for _uhr_raw, _cfg in UHR_BY_SCHOOL_CFG.items():
        _ovs_region = _cfg.get('overseas_region')
        if _ovs_region:
            _UHR_REGION_OVERRIDE.setdefault(_uhr_raw, _ovs_region)
    for _uhr_raw, _region in _UHR_REGION_OVERRIDE.items():
        uhr_region_map[normalize_uhr(_uhr_raw)] = _region
    uhr_rank = [
        {'rank': i, 'name': str(n), 'region': uhr_region_map.get(str(n), ''),
         'resumes': int(c), 'offers': int(uhr_offers.get(n, 0))}
        for i, (n, c) in enumerate(uhr_counts.items(), 1)
    ]

    # ---- 院校 ----
    school_counts = df['最高学历学校'].value_counts()
    school_offers = df[offer_mask].groupby('最高学历学校').size().to_dict()
    school_rank = [
        {'rank': i, 'name': str(n), 'region': _region_of(n) or '其他',
         'resumes': int(c), 'offers': int(school_offers.get(n, 0))}
        for i, (n, c) in enumerate(school_counts.items(), 1)
        if not pd.isna(n)
    ]  # 全部院校，不再截断

    # ---- 青云计划 ----
    is_qy = df['投递项目名称'].astype(str).str.contains('青云', na=False)
    qy_total = int(is_qy.sum())
    qy_intern = int(df[is_qy & (df['投递项目名称'] == '青云实习')].shape[0])
    qy_grad = int(df[is_qy & (df['投递项目名称'] == '青云计划-应届生')].shape[0])
    qy_offers = int((is_qy & _offer_bool).sum())
    qy_graduate_offer = int((is_qy & _offer_bool & (df['简历流程状态'] == '毕业生已录用')).sum())
    qy_intern_offer = int((is_qy & _offer_bool & (df['简历流程状态'] == '实习已录用')).sum())
    # 青云 TOP 岗位类
    qy_df = df[is_qy]
    qy_jobs = qy_df['投递岗位类'].value_counts().head(10)
    qy_jobs_list = [{'name': str(n), 'value': int(v)} for n, v in qy_jobs.items()]

    # 青云 Offer 明细
    qy_offer_df = df[is_qy & _offer_bool].copy()
    qy_offer_cols = [
        '校园大使名称', '校园大使学校', '校园大使部门', 'uhr名称', '学生姓名',
        '最高学历id', '最高学历名称', '最高学历学校', '最高学历学院', '最高学历毕业时间',
        '简历流程状态id', '简历流程状态', '投递岗位类'
    ]

    def _safe_str(v):
        try:
            if pd.isna(v):
                return ''
        except Exception:
            pass
        # 处理时间
        if hasattr(v, 'strftime'):
            try:
                return v.strftime('%Y-%m-%d')
            except Exception:
                return str(v)
        s = str(v).strip()
        if s.lower() in ('nan', 'nat', '<na>', 'none'):
            return ''
        # 浮点 id 去 .0
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return s

    qy_offer_details = []
    for _, row in qy_offer_df.iterrows():
        item = {}
        for c in qy_offer_cols:
            item[c] = _safe_str(row[c]) if c in qy_offer_df.columns else ''
        qy_offer_details.append(item)
    # 排序：状态 → 大使 → 学生
    qy_offer_details.sort(key=lambda r: (r.get('简历流程状态', ''),
                                          r.get('校园大使名称', ''),
                                          r.get('学生姓名', '')))

    # 青云 Offer TOP 排名（大使 / UHR / 课题=投递岗位类）
    def _qy_rank_by(col_label):
        """对 qy_offer_df 按指定列做 TOP 统计，返回 [{name, total, grad, intern}, ...] 全量倒序"""
        if col_label not in qy_offer_df.columns or qy_offer_df.empty:
            return []
        sub = qy_offer_df[[col_label, '简历流程状态']].copy()
        sub[col_label] = sub[col_label].apply(_safe_str)
        sub = sub[sub[col_label] != '']
        if sub.empty:
            return []
        agg = sub.groupby(col_label).agg(
            total=('简历流程状态', 'size'),
            grad=('简历流程状态', lambda s: (s == '毕业生已录用').sum()),
            intern=('简历流程状态', lambda s: (s == '实习已录用').sum()),
        ).reset_index().sort_values(['total', 'grad', 'intern'], ascending=[False, False, False])
        return [
            {'name': str(r[col_label]),
             'total': int(r['total']),
             'grad': int(r['grad']),
             'intern': int(r['intern'])}
            for _, r in agg.iterrows()
        ]

    qy_top_ambassador = _qy_rank_by('校园大使名称')
    qy_top_uhr = _qy_rank_by('uhr名称')
    qy_top_jobclass = _qy_rank_by('投递岗位类')

    # ---- 7个垂类 ----
    category_list = []
    for cat_name, keywords in CATEGORY_MAP.items():
        mask = pd.Series([False] * len(df))
        for kw in keywords:
            mask |= df['投递岗位类'].astype(str).str.contains(kw, na=False)
        cat_resumes = int(mask.sum())
        cat_offers = int((mask & _offer_bool).sum())
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
        'qingyunJobs': qy_jobs_list,
        'qingyunOfferDetails': qy_offer_details,
        'qingyunTopAmbassador': qy_top_ambassador,
        'qingyunTopUhr': qy_top_uhr,
        'qingyunTopJobclass': qy_top_jobclass,
        'regionOtherDetail': other_breakdown
    }
    return result


def build_last_year_payload():
    """从 LAST_YEAR_DATA 中只摘出对比展示需要的字段，瘦身后塞给前端。"""
    if not LAST_YEAR_DATA:
        return None
    s = LAST_YEAR_DATA['summary']
    # 排名表只保留 name → resumes/offers，前端按 name 匹配
    def pick(rows):
        return [{'name': r['name'], 'resumes': r['resumes'], 'offers': r['offers']} for r in rows]
    return {
        'summary': {
            'totalResumes': s['totalResumes'],
            'totalOffers': s['totalOffers'],
            'graduateOffer': s['graduateOffer'],
            'internOffer': s['internOffer'],
            'offerRate': s['offerRate'],
        },
        'regionRank': pick(LAST_YEAR_DATA['regionRank']),
        'uhrRank': pick(LAST_YEAR_DATA['uhrRank']),
        'ambassadorRank': pick(LAST_YEAR_DATA['ambassadorRank']),
        'schoolRank': pick(LAST_YEAR_DATA['schoolRank']),
    }


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

        # 优先用固定 sheet 名（兼容原逻辑），失败则用上面的智能匹配
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='获取名单上的这些外部伯乐的简历推荐数据_1')
        except Exception:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet)
        print(f"  Sheet: '{target_sheet}', 读取 {len(df)} 行")

        # ---- 读取 UHR 底表 ----
        df_uhr = load_uhr_df()

        # ---- 在覆盖之前，把"旧的当前数据"按 ISO 周备份成快照（同一周只保留首次）----
        try:
            take_weekly_snapshot(CURRENT_DATA)
        except Exception as _e:
            print(f"  ⚠️ 周快照生成失败（不影响本次上传）：{_e}")

        # ---- 处理数据 ----
        CURRENT_DATA = process_data(df, df_uhr)

        # ---- 持久化：避免 server 重启后丢失今年数据 ----
        try:
            with open(CURRENT_DATA_CACHE, 'wb') as f:
                pickle.dump(CURRENT_DATA, f)
            print(f"  💾 今年数据已持久化到 {os.path.basename(CURRENT_DATA_CACHE)}")
        except Exception as _e:
            print(f"  ⚠️ 今年数据持久化失败（不影响本次使用）：{_e}")

        t1 = time.time()
        print(f"  ✅ 处理完成，耗时 {t1-t0:.1f}s，{CURRENT_DATA['summary']['totalResumes']} 条简历")

        last_year = build_last_year_payload()
        return jsonify({
            'success': True,
            'message': f"导入成功！共 {CURRENT_DATA['summary']['totalResumes']:,} 条简历，{CURRENT_DATA['summary']['totalOffers']} 个 Offer",
            'data': {**CURRENT_DATA, 'lastYear': last_year},
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
    return jsonify({'loaded': True, **CURRENT_DATA, 'lastYear': build_last_year_payload()})


@app.route('/api/region_other')
def get_region_other():
    """诊断接口：列出全部落入「其他」区域的学校 + 命中条数 + 原因"""
    global CURRENT_DATA
    if CURRENT_DATA is None:
        return jsonify({'loaded': False, 'error': '请先上传内推数据 Excel'})
    detail = CURRENT_DATA.get('regionOtherDetail', [])
    return jsonify({
        'loaded': True,
        'total_schools': len(detail),
        'total_resumes': sum(x['count'] for x in detail),
        'detail': detail
    })


# ============================================================
# 周快照 / 周对比
#   - 每次 /api/upload 时，先把"覆盖前的旧 CURRENT_DATA"按 ISO 周编号备份一份到 snapshots/YYYY-Www.pkl
#   - 同一周内多次上传只保留最早那次（即"本周基线"），用于和最新数据对比"本周新增"
#   - /api/week_compare 返回：基准周 vs 当前数据 的核心指标 delta
# ============================================================
def _iso_week_key(ts=None):
    """返回 ISO 周键，例如 '2026-W21'。"""
    import datetime as _dt
    d = _dt.datetime.fromtimestamp(ts) if ts else _dt.datetime.now()
    y, w, _ = d.isocalendar()
    return f"{y}-W{w:02d}"


def _snapshot_path(week_key):
    return os.path.join(SNAPSHOT_DIR, f"{week_key}.pkl")


def take_weekly_snapshot(prev_data):
    """
    把"覆盖前的旧 CURRENT_DATA"备份成本周快照。
    同一周内只保留最早那一次（即"本周基线"），后续上传不会覆盖它。
    """
    if prev_data is None:
        return None
    week_key = _iso_week_key()
    path = _snapshot_path(week_key)
    if os.path.exists(path):
        return week_key  # 本周已有基线，不覆盖
    try:
        with open(path, 'wb') as f:
            pickle.dump({
                'week_key': week_key,
                'snapshot_time': time.time(),
                'summary': prev_data.get('summary'),
                'ambassadorRank': prev_data.get('ambassadorRank'),
                'regionRank': prev_data.get('regionRank'),
                'uhrRank': prev_data.get('uhrRank'),
            }, f)
        print(f"  📸 已生成周快照 {week_key}.pkl（基线= {prev_data['summary']['totalResumes']} 简历 / {prev_data['summary']['totalOffers']} Offer / {prev_data['summary']['eliteCount']} 精英大使）")
        return week_key
    except Exception as e:
        print(f"  ⚠️ 周快照保存失败：{e}")
        return None


def list_snapshots():
    """列出所有周快照，按时间倒序。返回 [{week_key, snapshot_time, summary}]"""
    items = []
    if not os.path.isdir(SNAPSHOT_DIR):
        return items
    for fn in os.listdir(SNAPSHOT_DIR):
        if not fn.endswith('.pkl'):
            continue
        path = os.path.join(SNAPSHOT_DIR, fn)
        try:
            with open(path, 'rb') as f:
                d = pickle.load(f)
            items.append({
                'week_key': d.get('week_key') or fn.replace('.pkl', ''),
                'snapshot_time': d.get('snapshot_time'),
                'summary': d.get('summary'),
                'file': fn,
            })
        except Exception:
            continue
    items.sort(key=lambda x: x.get('snapshot_time') or 0, reverse=True)
    return items


def _load_snapshot(week_key):
    path = _snapshot_path(week_key)
    if not os.path.exists(path):
        return None
    try:
        with open(path, 'rb') as f:
            return pickle.load(f)
    except Exception:
        return None


def _build_amb_diff(curr_rank, prev_rank, threshold=ELITE_THRESHOLD):
    """
    对比"当前 vs 基线"的精英大使变化。
    返回：{newElites: [...], lostElites: [...], summary: {curr_count, prev_count, delta_count}}
    """
    curr_map = {a['name']: a for a in (curr_rank or [])}
    prev_map = {a['name']: a for a in (prev_rank or [])}
    curr_elites = {n for n, a in curr_map.items() if a.get('resumes', 0) >= threshold}
    prev_elites = {n for n, a in prev_map.items() if a.get('resumes', 0) >= threshold}

    new_names = sorted(curr_elites - prev_elites)
    lost_names = sorted(prev_elites - curr_elites)

    new_list = []
    for n in new_names:
        c = curr_map.get(n, {})
        p = prev_map.get(n, {})
        new_list.append({
            'name': n,
            'currResumes': c.get('resumes', 0),
            'prevResumes': p.get('resumes', 0),
            'addResumes': c.get('resumes', 0) - p.get('resumes', 0),
            'currOffers': c.get('offers', 0),
            'prevOffers': p.get('offers', 0),
        })
    lost_list = []
    for n in lost_names:
        c = curr_map.get(n, {})
        p = prev_map.get(n, {})
        lost_list.append({
            'name': n,
            'currResumes': c.get('resumes', 0),
            'prevResumes': p.get('resumes', 0),
        })

    # 简历增长 TOP 大使（不限定精英身份）
    growth = []
    for n, c in curr_map.items():
        p = prev_map.get(n, {})
        d = c.get('resumes', 0) - p.get('resumes', 0)
        if d > 0:
            growth.append({
                'name': n,
                'add': d,
                'currResumes': c.get('resumes', 0),
                'prevResumes': p.get('resumes', 0),
                'currOffers': c.get('offers', 0),
                'prevOffers': p.get('offers', 0),
                'addOffers': c.get('offers', 0) - p.get('offers', 0),
                'wasElite': p.get('resumes', 0) >= threshold,
                'isElite': c.get('resumes', 0) >= threshold,
            })
    growth.sort(key=lambda x: x['add'], reverse=True)

    return {
        'newElites': new_list,
        'lostElites': lost_list,
        'eliteCurrCount': len(curr_elites),
        'elitePrevCount': len(prev_elites),
        'eliteDelta': len(curr_elites) - len(prev_elites),
        'growthTop': growth[:30],
    }


@app.route('/api/week_compare')
def get_week_compare():
    """
    返回"基线周 vs 当前数据"的对比。
    可选参数 ?week=YYYY-Www 指定对比哪一周；不传则取最近一次"非本周"的快照（即"上一周基线"）。
    """
    global CURRENT_DATA
    if CURRENT_DATA is None:
        return jsonify({'loaded': False, 'error': '请先上传内推数据 Excel'})

    snaps = list_snapshots()
    if not snaps:
        return jsonify({
            'loaded': True,
            'has_baseline': False,
            'message': '暂无周快照基线 — 下次上传 Excel 时会自动建立基线，再下次上传后即可看到周对比',
        })

    target = request.args.get('week')
    this_week = _iso_week_key()
    if target:
        snap = _load_snapshot(target)
    else:
        # 优先取"上一周"的快照；若只有本周快照（说明这是首次有基线），就用本周基线
        non_this = [s for s in snaps if s['week_key'] != this_week]
        pick_key = (non_this[0]['week_key'] if non_this else snaps[0]['week_key'])
        snap = _load_snapshot(pick_key)

    if not snap:
        return jsonify({
            'loaded': True,
            'has_baseline': False,
            'message': '未找到指定周快照',
            'available_weeks': [s['week_key'] for s in snaps],
        })

    s_curr = CURRENT_DATA['summary']
    s_prev = snap.get('summary') or {}

    def diff_num(k):
        c = s_curr.get(k, 0) or 0
        p = s_prev.get(k, 0) or 0
        return {'curr': c, 'prev': p, 'delta': c - p}

    amb_diff = _build_amb_diff(
        CURRENT_DATA.get('ambassadorRank'),
        snap.get('ambassadorRank'),
    )

    return jsonify({
        'loaded': True,
        'has_baseline': True,
        'baseline_week': snap.get('week_key'),
        'baseline_time': snap.get('snapshot_time'),
        'this_week': this_week,
        'available_weeks': [{'week_key': s['week_key'], 'snapshot_time': s['snapshot_time']} for s in snaps],
        'kpi': {
            'totalResumes': diff_num('totalResumes'),
            'totalOffers': diff_num('totalOffers'),
            'totalAmbassadors': diff_num('totalAmbassadors'),
            'totalSchools': diff_num('totalSchools'),
            'eliteCount': diff_num('eliteCount'),
            'qingyunResumes': diff_num('qingyunResumes'),
            'qingyunOffers': diff_num('qingyunOffers'),
            'graduateOffer': diff_num('graduateOffer'),
            'internOffer': diff_num('internOffer'),
        },
        'ambassadors': amb_diff,
    })


@app.route('/api/snapshots')
def get_snapshots():
    """诊断用：列出所有周快照"""
    return jsonify({'snapshots': list_snapshots()})


# ============================================================
# 启动前预加载
# ============================================================
def _bootstrap():
    print("=" * 50)
    print("🚀 校园大使数据看板 - 启动初始化")
    print(f"   UHR_FILE              = {UHR_FILE}  (exists={os.path.exists(UHR_FILE)})")
    print(f"   LAST_YEAR_FILE        = {LAST_YEAR_FILE}  (exists={os.path.exists(LAST_YEAR_FILE)})")
    print(f"   CATEGORY_BASELINE_FILE= {CATEGORY_BASELINE_FILE}  (exists={os.path.exists(CATEGORY_BASELINE_FILE)})")
    load_category_map()
    load_last_year_data()

    # 命中今年数据缓存：避免 server 重启后丢失上一次上传
    global CURRENT_DATA
    try:
        if os.path.exists(CURRENT_DATA_CACHE):
            with open(CURRENT_DATA_CACHE, 'rb') as f:
                CURRENT_DATA = pickle.load(f)
            print(f"  ⚡ 命中今年数据缓存：简历 {CURRENT_DATA['summary']['totalResumes']} / Offer {CURRENT_DATA['summary']['totalOffers']}")
    except Exception as _e:
        print(f"  ⚠️ 今年数据缓存加载失败，等待重新上传：{_e}")
        CURRENT_DATA = None
    print("=" * 50)


_bootstrap()


# ============================================================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8765))
    print(f"   监听端口: {port}")
    print("   上传 Excel 后即可生成看板")
    print("=" * 50)
    app.run(host='0.0.0.0', port=port, debug=False)
