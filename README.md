# 校园大使内推数据看板（Campus Ambassador Dashboard）

> Flask + ECharts 单页看板，把内推 Excel 跑成多维交互式可视化大屏，并提供「今年 vs 去年」两年对比专区。

![status](https://img.shields.io/badge/status-active-success)
![python](https://img.shields.io/badge/python-3.8%2B-blue)
![flask](https://img.shields.io/badge/flask-3.0.3-lightgrey)

---

## ✨ 主要特性

- **一文件全栈**：后端单 `server.py` 200+ 行 process_data / 前端单 `dashboard.html` ~3000 行（含全部 ECharts 渲染）。
- **拖拽上传 Excel → 实时出图**：自动识别「外部伯乐 / 简历推荐」sheet。
- **8 张顶部 KPI**：入库简历、Offer、Offer 率、精英大使、青云计划、人均推荐、青云 Offer、7 个垂类简历。
- **多维下钻**：区域 / UHR（带筛选 pills）/ 学校 / 校园大使 / 7 大业务垂类 / 简历状态分布。
- **🚀 青云计划专区**：TOP 10 投递岗位类、3 张 TOP 榜（大使 / UHR / 岗位类）、明细表三重筛选。
- **📅 两年对比专区**（重点功能）：
  - 总览四卡（简历 / Offer / Offer 率 / UHR 数）
  - 区域今年 vs 去年
  - 院校 TOP 15 内推数对比
  - **👥 UHR 全员两年对比**（横向条形 + 9 列明细表，新进入打 `[新]`、今年=0 打 `[今年无]`）
- **周快照**：`set_baseline.py` 把任意历史 Excel 烤成 ISO 周快照，看板「本周 vs 上周」自动出。
- **目录监听**：`watch_folder.py` 监控指定文件夹，新 Excel 落盘即自动刷新 JSON 数据。

---

## 🧱 技术栈

| 层 | 选型 |
|---|---|
| 后端 | Python 3.8+ · Flask 3 · pandas 2.2 · openpyxl 3.1 · gunicorn 22 |
| 前端 | 原生 HTML + ES2017 + ECharts 5（本仓自带 `echarts.min.js`，离线可用） |
| 数据 | Excel（xlsx）· pickle 缓存（不入库） |
| 部署 | gunicorn 单进程多线程 / Procfile / Docker / 任意 Linux 云主机 |

---

## 🗂️ 目录结构

```
campus-ambassador-dashboard/
├── server.py                       # Flask 入口 + 全部 process_data 数据处理
├── dashboard.html                  # 单文件前端看板（ECharts）
├── echarts.min.js                  # 离线版 ECharts 5
├── watch_folder.py                 # 文件夹监听器（新 xlsx 自动跑数据）
├── set_baseline.py                 # 把历史 Excel 烤成 ISO 周快照
├── check_elite_diff.py             # 临界区精英大使比对脚本
├── requirements.txt
├── Procfile                        # web: gunicorn ...
├── start.sh                        # 一键安装依赖 + 启动脚本
├── uhr_overrides.example.json      # UHR 区域 override 配置模板
├── category_map.example.json       # 垂类岗位关键词模板
├── category_dept_map.example.json  # baseline 部门→垂类映射模板
├── .gitignore
└── README.md
```

> ⚠️ 仓库**不含**任何业务数据：`*.xlsx / *.csv / *.pkl / snapshots/ / *.log` 全部已加入 `.gitignore`。
> 业务方真实使用的 `uhr_overrides.json` / `category_map.json` / `category_dept_map.json`（含人名/业务线名）也在 `.gitignore` 里。
> 真实数据请放在本地或私有存储，勿提交。

---

## 🚀 本地快速启动

```bash
# 1. 装依赖
pip install -r requirements.txt

# 2. 启动
python server.py
# → http://localhost:8765
```

打开后页面会先显示空状态；点页面里**「上传 Excel」**按钮选一个内推数据 xlsx，约 2~10 秒就能看到全部图表。

### 配置项（环境变量）

| 变量 | 说明 | 默认 |
|---|---|---|
| `PORT` | 监听端口 | `8765` |
| `UHR_FILE` | UHR-高校映射底表绝对路径（学校 → UHR / 区域） | 同目录 `UHR-高校底表.xlsx`；不存在则启动会提示 |
| `LAST_YEAR_FILE` | 去年权威数据 xlsx，用于「两年对比」专区 | 不设置则跳过去年模块 |
| `CATEGORY_BASELINE_FILE` | 历史垂类基线 xlsx | 不设置则用 JSON 配置 |
| `UHR_OVERRIDES_FILE` | UHR 区域 override JSON 路径 | 同目录 `uhr_overrides.json`；缺省则不做硬覆盖 |
| `CATEGORY_MAP_FILE` | 垂类岗位关键词 JSON | 同目录 `category_map.json`；缺省则空字典 |
| `CATEGORY_DEPT_MAP_FILE` | baseline 部门 → 垂类卡片名 JSON | 同目录 `category_dept_map.json`；缺省则空字典 |
| `WATCH_DIR` | `watch_folder.py` 监听目录 | `~/Desktop` |

### 业务配置 JSON（含敏感名单，不入库）

仓库提交了 3 份 `*.example.json` 模板，**真实文件命名（去掉 `.example`）已加入 `.gitignore`**，由部署方在本机准备：

| 模板文件 | 作用 |
|---|---|
| `uhr_overrides.example.json` | UHR 主辖区域硬覆盖 / 跨区 UHR 按学校粒度归属配置 |
| `category_map.example.json` | 垂类卡片名 → 岗位类关键词数组 |
| `category_dept_map.example.json` | baseline xlsx 中「部门」字段 → 垂类卡片名 |

```bash
# 用法：复制模板 → 改成真实业务数据
cp uhr_overrides.example.json uhr_overrides.json
cp category_map.example.json category_map.json
cp category_dept_map.example.json category_dept_map.json
# 然后按本机业务实际填名单 / 关键词；server.py 启动时会自动加载
```

### 可选：启动文件夹监听（新数据自动入库）

```bash
# 默认监听 ~/Desktop 下的 *.xlsx；可用 WATCH_DIR 环境变量改
python watch_folder.py
```

把新 Excel 拖到桌面就会自动 `process_data` → 写 `dashboard_data.json` → 归档到 `processed/`。

---

## 🐳 云端部署

### 方案 A：Procfile（Heroku / 任意 PaaS）

```
web: gunicorn -w 1 --threads 4 -b 0.0.0.0:${PORT:-8765} --timeout 120 server:app
```

### 方案 B：start.sh（Linux 云主机 / 容器）

```bash
chmod +x start.sh
./start.sh
```

### 方案 C：Docker（自行写一份 Dockerfile）

```dockerfile
FROM python:3.10-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV PORT=8765
CMD ["gunicorn","-w","1","--threads","4","-b","0.0.0.0:8765","--timeout","120","server:app"]
```

---

## 📊 数据规范

### 内推主表（每周更新的那个 xlsx）

至少需要以下列（顺序无所谓）：

| 列名 | 示例 | 说明 |
|---|---|---|
| 大使姓名 | （字符串） | 校园大使 |
| 外部伯乐对接uhr | （字符串） | UHR 名字会用 `normalize_uhr()` 抹平全/半角括号 |
| 投递岗位类 | 后台开发 / 算法 / ... | 7 大垂类聚合用 |
| 简历流程状态 | 已 Offer / 面试中 / ... | Offer 数 / 状态分布用 |
| 最高学历学校 | 清华大学 / Stanford 等 | 院校排名 + 跨区 UHR 区域细分用 |
| 是否青云计划 | 是 / 否 | 青云专区筛选 |
| ... | | 其它字段（联系方式等）一律不入库，只走内存 |

### UHR-高校底表 (`UHR_FILE`)

`Sheet2` 必填，至少 3 列：`学校名 / UHR / 区域`。区域白名单 11 个：

```
华北 / 华东 / 华南 / 华中 / 西区 / 港校 / 欧洲 / 北美 / 亚太 / OHR / IEG
```

### 去年权威数据（用于两年对比）

可选，结构与今年表一致，关键 sheet 命名约定：
- 「底表（全）」/「底表（offer）」 — 全量明细
- 「洗数据（排名）」 — 大使 → UHR 映射（13 个 UHR）
- 「洗数据（offer）」 — 去年 Offer 总数核对

---

## 🧠 核心算法 / 业务规则

### 1. UHR 区域映射

- 默认按「该 UHR 名下学校最多归属到哪个区域」自动建字典；
- 通过 `UHR_REGION_OVERRIDE` 字典做兜底（少数跨区 UHR 显式指定）；
- **特殊处理**：对于名下学校横跨海外+大陆的 UHR，按**最高学历学校粒度**严格区分：
  - 海外院校（含港澳台/澳大利亚等）→ **亚太**
  - 中国大陆院校 → 该 UHR 在大陆的主区域（如 **华北**）
  - 实现：`_is_overseas_or_hkmotw_school()` + `_HKMOTW_KEYWORDS` + `_MAINLAND_OVERRIDE_KEYWORDS`（清华/北京/宁波诺丁汉/昆山杜克/西交利物浦/香港中文大学（深圳）等大陆办学的港校分校）。
  - 具体哪些 UHR 走该特判由 `UHR_REGION_OVERRIDE` 配置，本仓不带名单。

### 2. 精英大使

- 内推数 ≥ 200 即为精英；周对比时通过 `snapshots/<ISO-week>.pkl` 拿基线。

### 3. 青云计划三榜

`server.py` 内部 `_qy_rank_by(col_label)` 同一函数生成「校园大使 / UHR / 投递岗位类」TOP 榜。

### 4. UHR 全员两年对比

前端 `renderCmpUHR()` 取「今年 uhrRank ∪ 去年 uhrRank」并集去重；按今年简历降序，今年=0 的按去年简历降序排尾。新出现的 UHR 标 `[新]`，去年存在但今年=0 标 `[今年无]`。

---

## 🔌 主要接口

| Path | Method | 说明 |
|---|---|---|
| `/` | GET | dashboard.html |
| `/api/upload` | POST | 上传 Excel，触发 `process_data`，返回完整看板 JSON |
| `/api/data` | GET | 读当前内存里的数据（含 lastYear 子树） |
| `/api/week_compare` | GET | 本周 vs 上周（自动选最近一份非本周快照） |
| `/static/echarts.min.js` | GET | 离线 ECharts |

---

## 🧪 开发提示

- 修改 `process_data / load_last_year_data / UHR_REGION_OVERRIDE / _is_overseas_or_hkmotw_school` 后，**务必删掉 `last_year_data.pkl` 让其重算**，不然你看到的还是旧数据。
- Windows 命令行临时跑校验脚本：`set PYTHONIOENCODING=utf-8 && set PYTHONUTF8=1 && python check_xxx.py`，避免 emoji / 中文乱码。
- 前端是单文件 ~3000 行原生 HTML，**没用任何打包工具**，直接 Ctrl+F 搜函数名定位。

---

## 📜 License

本项目仅作内部数据可视化工具示例。代码部分以 MIT 协议开源；任何业务数据（简历 / Offer / UHR 名单 / 学校映射）均不在仓库内，使用方需自行准备。

---

## 🙏 致谢

- [Apache ECharts](https://echarts.apache.org/) — 全部图表
- [Flask](https://flask.palletsprojects.com/) — 后端
- [pandas](https://pandas.pydata.org/) — 数据处理

如果这个项目对你有用，欢迎 ⭐！
