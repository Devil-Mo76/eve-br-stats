# EVE Online BR 战斗统计系统（通用版）

> 一个用于分析 **EVE Online** 战斗报告（BR）的本地统计工具：粘贴 BR 链接，自动汇总参战成员的
> **出击次数 / 击杀数 / 总伤害 / 使用舰船**，导出带柱状图的 Excel 报表；同时提供舰队**出勤统计**功能。
> 本仓库为**通用版本**，不含任何特定军团、联盟或个人的数据与标识。
>
> A local EVE Online battle-report analyser: paste BR links, get per-pilot **sorties / kills /
> damage / ships flown** aggregated into an Excel report with charts, plus a fleet
> **attendance tracker**. This is the **generic edition** — no corporation, alliance or
> personal data is bundled.

**🌐 文档 / Documentation**
[中文操作手册](#中文操作手册) · [English Operation Manual](#english-operation-manual) ·
[操作手册.md](操作手册.md) · [OPERATION_MANUAL_EN.md](OPERATION_MANUAL_EN.md) ·
[数据文件格式](data/README.md)

## 快速开始 / Quick start

```bash
pip install -r requirements.txt      # 安装依赖 / install dependencies
python app.py                        # 启动 / start (http://127.0.0.1:5000)
```

首次使用需在网页 **管理 → 军团管理** 添加你要统计的军团 ID；
其余配置集中在 [`config.py`](config.py)，源码中搜索 `TODO(通用版)` 或 `【需修改】` 即可定位全部待改项。

On first run, add the corporation you want to track under **管理 → 军团管理** in the web UI.
Everything else is configured in [`config.py`](config.py); search the source for
`TODO(通用版)` or `【需修改】` to find every option you may want to change.

| 模块 / Module | 说明 / Description |
| --- | --- |
| 战斗统计分析 | 解析多个 BR 链接，按军团筛选成员，统计出击/击杀/伤害/舰船 |
| 出勤统计分析 | 按活动逐次录入舰队成员，实时统计出勤次数并导出 Excel |
| 报表导出 | Excel 含数据表 + 「总伤害 Top10」「击杀数 Top10」两张柱状图 |
| 管理后台 | 军团管理 / 角色管理 / 历史报告 / 系统日志 |

其他命令 / Other commands：

```bash
python tools/smoke_test.py     # 冒烟测试（25 项断言，数据目录自动隔离）
python ship_id.py              # 更新舰船 ID→名称 映射（报表出现 Unknown 时用）
python build.py                # 打包 dist/BRStatisticsTool.exe
```

---

## 中文操作手册

## EVE Online BR 战斗统计系统 · 详细操作手册

> **中文版** ｜ English version → [OPERATION_MANUAL_EN.md](#english-operation-manual)
> 适用版本：v1.1 通用版
> 适用读者：部署者（第一次接触本程序的人）、日常使用者、二次开发者

---

### 阅读指引

| 你的身份 | 建议阅读 |
| --- | --- |
| 第一次拿到这个程序，想赶紧用起来 | 第一部分 → 第二部分第 6~9 章 → 第三部分 |
| 要改品牌、改配置、改端口 | 第二部分第 10~11 章 |
| 要打包成 EXE 发给朋友 | 第四部分 |
| 用着用着出错了 | 第五部分第 21~22 章 |
| 要改代码、加功能 | 第六部分 |

全文出现的 `[需修改]` 标记，对应源码里真实存在的 `TODO(通用版)` / `【需修改】` 注释，
可直接搜索定位。

---

### 目录

**第一部分 · 认识系统**
- [1. 系统是什么](#1-系统是什么)
- [2. 功能清单](#2-功能清单)
- [3. 工作原理与数据流](#3-工作原理与数据流)
- [4. 目录结构](#4-目录结构)

**第二部分 · 部署与配置**
- [5. 环境要求](#5-环境要求)
- [6. 安装（源码方式）](#6-安装源码方式)
- [7. 启动与停止](#7-启动与停止)
- [8. 三种部署形态](#8-三种部署形态)
- [9. 首次必做配置](#9-首次必做配置)
- [10. 配置项完整参考（config.py）](#10-配置项完整参考configpy)
- [11. 需自行修改清单](#11-需自行修改清单)

**第三部分 · 功能详解**
- [12. 战斗统计分析](#12-战斗统计分析)
- [13. 出勤统计分析](#13-出勤统计分析)
- [14. 管理后台](#14-管理后台)
- [15. 报表文件说明](#15-报表文件说明)

**第四部分 · 打包与发布**
- [16. 打包为 EXE](#16-打包为-exe)
- [17. 分发与迁移](#17-分发与迁移)
- [18. 上传到 GitHub](#18-上传到-github)

**第五部分 · 运维与排错**
- [19. 数据备份与恢复](#19-数据备份与恢复)
- [20. 升级流程](#20-升级流程)
- [21. 错误信息速查表](#21-错误信息速查表)
- [22. 常见问题 FAQ](#22-常见问题-faq)

**第六部分 · 开发者指引**
- [23. 代码结构与关键函数](#23-代码结构与关键函数)
- [24. 二次开发指引](#24-二次开发指引)
- [25. 通用版变更记录](#25-通用版变更记录)

**附录**
- [A. 数据文件格式](#附录-a数据文件格式)
- [B. 配置项速查表](#附录-b配置项速查表)
- [C. 术语表](#附录-c术语表)

---

## 第一部分 · 认识系统

### 1. 系统是什么

**EVE Online BR 战斗统计系统** 是一个跑在你自己电脑上的本地网页工具。

它的核心工作是：**把一串 BR（Battle Report，战斗报告）链接，翻译成一张「谁打了多少伤害」的排行榜。**

EVE Online 的战斗记录分散在第三方站点上，一场会战可能产生几十上百个 BR 链接。
手工统计谁出力多、谁划水，基本不可行。本系统把这件事自动化了：

```
你粘贴一堆 BR 链接  →  程序逐个抓取战斗记录  →  只挑出你指定军团的成员  →  汇总成 Excel
```

**它不是**：不是 EVE 官方工具，不修改游戏，不读取你的账号，不需要登录。
它只做一件事——**公开数据聚合**。

### 2. 功能清单

| # | 功能 | 入口 | 输出 |
| --- | --- | --- | --- |
| 1 | 战斗统计 | 首页「战斗统计分析」 | 带图表的 Excel（排名/角色/出击/击杀/伤害/舰船） |
| 2 | 出勤统计 | 导航栏「出勤统计分析」 | 出勤次数排行 + Excel（含历史记录表） |
| 3 | 军团管理 | 管理 → 军团管理 | 维护「要统计哪些军团」 |
| 4 | 角色管理 | 管理 → 角色管理 | 维护角色 ID↔名称映射 |
| 5 | 历史报告 | 管理 → 历史报告 | 查看/下载/删除历史报表 |
| 6 | 系统日志 | 管理 → 系统日志 | 排错依据 |

### 3. 工作原理与数据流

#### 3.1 战斗统计流程

```
 ① 用户粘贴 BR 链接
        │
        ▼
 ② 格式校验（正则，前后端各校一次）
        │  不合规 → 页面红字提示，不发起请求
        ▼
 ③ GET https://br.evetools.org/api/v1/composition/get/{BR_ID}
        │  取到这一场战斗的 relateds → kms[]（击杀邮件列表）
        ▼
 ④ 对每条击杀邮件：GET https://kb.evetools.org/api/v1/killmails/{KILL_ID}
        │
        ├── 提取 atts[]（攻击者列表）与 victim（受害者）
        ├── 用 corp（军团ID）比对 data/corp_id.txt
        │      ├─ 命中 → 累加该角色的出击次数 / 伤害
        │      └─ 未命中 → 直接跳过（这是"只统计自己人"的关键）
        ├── 若攻击者 blow（final blow）= true → 该角色击杀数 +1
        └── 遇到没见过的角色 → 自动写入 data/char_id.txt 补全名称
        ▼
 ⑤ generate_final_data() 汇总：出击次数 / 击杀数 / 总伤害 / 使用舰船（去重）
        ▼
 ⑥ openpyxl 生成 Excel：数据表 + 伤害Top10柱状图 + 击杀Top10柱状图
        ▼
 ⑦ 存到 data/reports/，页面出现「下载最新报告」按钮
```

**关键概念说明**

| 概念 | 含义 |
| --- | --- |
| 出击次数（Sorties） | 该角色在本次统计的所有 BR 中出现的次数（同一场只计一次） |
| 击杀数（Kills） | 打出**致命一击**（final blow）的次数，不是参与次数 |
| 总伤害（Damage） | 所有击杀邮件里该角色造成伤害的累加 |
| 使用舰船（Ships） | 期间使用过的舰船名去重后的列表 |

> ⚠️ **重要**：统计范围由 `data/corp_id.txt` 决定。如果这里为空，
> 程序会直接报错并引导你去配置——**这是刻意设计**，避免生成一份全是空白的报表让你摸不着头脑。

#### 3.2 出勤统计流程

```
 ① 填写日期 + 活动类型 + 成员名单（每行一个）
        │
        ▼
 ② 前端 fetch POST → /attendance/submit_fleet
        │
        ├── 按顺序去重（同一份名单里重复的角色只算一次）
        └── 追加写入 data/attendance.json
        ▼
 ③ 前端立即 GET /attendance/get_stats 刷新右侧实时表格
        ▼
 ④ 点「导出Excel」→ /attendance/export_stats
        └── 内存中生成 xlsx（工作表：出勤统计 / 历史记录）直接下载，不落临时文件
```

> 出勤数据**持久化到磁盘**，关掉程序再打开，数据还在。

#### 3.3 涉及的第三方接口

| 用途 | 地址 | 说明 |
| --- | --- | --- |
| BR 详情 | `https://br.evetools.org/api/v1/composition/get/{id}` | 取战斗组成的击杀邮件列表 |
| 击杀邮件 | `https://kb.evetools.org/api/v1/killmails/{id}` | 取攻击者/受害者明细 |
| 舰船数据 | `https://esi.evetech.net/latest` | 官方 ESI，仅 `ship_id.py` 使用 |

以上地址都写在 `config.py` 中，接口变更时可自行替换。

### 4. 目录结构

```
br-web/
│
├── app.py                      # 主程序：路由 + 战斗统计核心 + Excel 生成
├── config.py                   # ★ 唯一配置入口（品牌/端口/接口/出勤选项）
├── attendance.py               # 出勤统计蓝图（含 JSON 持久化）
├── ship_id.py                  # 舰船 ID→名称 更新脚本
├── build.py                    # PyInstaller 打包脚本
├── version_info.txt            # EXE 版本信息（公司名/版权）
├── requirements.txt            # 依赖清单
│
├── 操作手册.md                 # ★ 本文件（中文详细版）
├── OPERATION_MANUAL_EN.md      # ★ English full manual
├── README.md                   # 项目简介
├── app_icon.ico                # EXE 图标
│
├── tools/
│   ├── make_icon.py            # 生成 favicon / exe 图标 / PNG LOGO
│   └── smoke_test.py           # 冒烟测试（25 项断言，数据目录自动隔离）
│
├── data/                       # 运行时数据（详见 data/README.md）
│   ├── ship_id.txt             # 舰船映射（公共数据，已预置 6000+ 条）
│   ├── corp_id.txt             # 军团配置 ← 首次使用要填这里
│   ├── char_id.txt             # 角色映射（统计时自动写入）
│   ├── br_links.txt            # 上次提交的 BR 链接（自动保存）
│   ├── attendance.json         # 出勤记录（自动生成）
│   ├── reports/                # 生成的 Excel 报表
│   └── README.md
│
├── static/
│   ├── site_logo.svg           # 导航栏 LOGO（中性占位，可替换）
│   ├── site_logo.png           # 同上，PNG 版
│   ├── favicon.ico             # 浏览器标签图标
│   └── style.css
│
└── templates/                  # Jinja2 页面模板
    ├── base.html               # 母版：导航、页脚、背景、样式
    ├── index.html              # 战斗统计首页
    ├── attendance.html         # 出勤统计页
    ├── docs.html               # 内置操作文档页
    ├── manage_corps.html       # 军团管理
    ├── manage_chars.html       # 角色管理
    ├── history.html            # 历史报告
    └── logs.html               # 系统日志
```

---

## 第二部分 · 部署与配置

### 5. 环境要求

| 项目 | 最低要求 | 推荐 |
| --- | --- | --- |
| 操作系统 | Windows 10 / 11、Linux、macOS | Windows 11 |
| Python | 3.9 | 3.10 ~ 3.12 |
| 内存 | 512 MB | 2 GB |
| 磁盘 | 200 MB（含舰船数据） | 1 GB（留报表空间） |
| 网络 | 能访问 `br.evetools.org`、`kb.evetools.org` | 直连，**关闭 VPN/代理** |
| 浏览器 | Chrome / Edge / Firefox 任一现代浏览器 | Chrome / Edge |

**关于 VPN**：这一步很关键。BR 接口对频繁请求有速率限制，
挂代理容易触发限流或被判定异常，官方页面也提示「尽量不要开启 VPN」。

### 6. 安装（源码方式）

```bash
# 1) 进入项目目录
cd br-web

# 2) 创建虚拟环境（推荐，避免污染系统 Python）
python -m venv .venv

# 3) 激活虚拟环境
.venv\Scripts\activate           # Windows
# source .venv/bin/activate      # Linux / macOS

# 4) 安装依赖
pip install -r requirements.txt

# 5) 启动
python app.py
```

**国内网络 pip 慢**，可换镜像源：

```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

依赖说明：

| 包 | 用途 | 必需 |
| --- | --- | --- |
| Flask | Web 框架 | ✅ |
| requests | 调用第三方接口 | ✅ |
| openpyxl | 生成 Excel 与图表 | ✅ |
| pandas | 出勤 Excel 导出 | ✅ |
| pyinstaller | 打包 EXE | 仅打包时需要 |
| pillow | 生成图标 | 仅改图标时需要 |

### 7. 启动与停止

```bash
python app.py
```

- 启动后会**自动打开浏览器**（由 `config.AUTO_OPEN_BROWSER` 控制）；
- 控制台会打印数据目录和访问地址，**不要关这个窗口**；
- 停止：在该窗口按 `Ctrl + C`，或直接关闭窗口。

**端口被占用**时改 `config.py` 的 `PORT`（5001、8000 均可）。
查看谁占用了 5000：

```bash
netstat -ano | findstr :5000     # Windows
lsof -i :5000                    # Linux / macOS
```

### 8. 三种部署形态

#### 形态一：本机自用（默认，最省事）

```python
# config.py
HOST = "127.0.0.1"
PORT = 5000
```

只有自己能访问，无需考虑安全问题。

#### 形态二：局域网共用（给队友用）

```python
# config.py
HOST = "0.0.0.0"           # 监听所有网卡
PORT = 5000
SECRET_KEY = "换成你自己的随机字符串"
```

其他人访问 `http://<你的内网IP>:5000`。
查自己的内网 IP：Windows 执行 `ipconfig`，找「IPv4 地址」。
**还需要在 Windows 防火墙放行该端口**（首次运行 Python 时点击「允许访问」通常即可）。

#### 形态三：公网服务器（谨慎）

仅在你确实需要时才做，且必须：

1. 修改 `SECRET_KEY` 为随机字符串；
2. 用 Nginx 反向代理 + HTTPS；
3. 本程序**没有用户认证**，公网暴露意味着任何人都能生成报表、查看历史。
   如需认证，见[第 24.4 节](#244-想加登录认证怎么做)自行扩展。

### 9. 首次必做配置

程序装好后，**只差一步就能用**：告诉它统计哪个军团。

#### 9.1 通过网页配置（推荐）

1. 打开 <http://127.0.0.1:5000>
2. 顶部导航 → **管理 → 军团管理**
3. 填写：
   - **军团ID**：纯数字，例如 `98769610`
   - **军团名称**：仅用于展示，随便写
4. 点「添加军团」

#### 9.2 军团 ID 怎么找

打开 <https://zkillboard.com> → 搜索你的军团 → 进入军团主页，看地址栏：

```
https://zkillboard.com/corporation/98769610/
                                └────┬───┘
                                  这就是军团 ID
```

也可以从 BR 工具 <https://br.evetools.org> 的军团页里取。

#### 9.3 直接编辑文件（可选）

也可以直接编辑 `data/corp_id.txt`，每行 `ID_名称`：

```
98769610_Example Corporation
12345678_Another Corporation
```

空行和 `#` 开头的注释行会被忽略。

#### 9.4 验证是否配置成功

回到首页粘贴任意一个 BR 链接生成报告。若未配置军团，页面会红字提示
「尚未配置任何军团 ID，请先到「军团管理」页面添加你要统计的军团」——
看到这句话就说明军团配置还是空的。

### 10. 配置项完整参考（config.py）

所有可配置项都在 `config.py`，下面逐项说明。**标 ★ 的是通常需要你改的。**

#### 10.1 品牌信息

| 变量 | 默认值 | 说明 |
| --- | --- | --- |
| ★ `SITE_NAME` | `EVE Online BR 战斗统计系统` | 浏览器标题、导航栏、页脚统一使用这个名称 |
| ★ `SITE_VERSION` | `1.1` | 显示为「系统名 v1.1」 |
| ★ `FOOTER_TEXT` | 同 `SITE_NAME` | 页脚文案，最终显示「页脚文案 ©年份」 |
| ★ `CORP_LOGO` | `site_logo.svg` | 导航栏 LOGO。填 `""` 则不显示 LOGO |
| `BACKGROUND_IMAGE` | `""` | 页面背景图。留空使用内置深空渐变 |

**换成自己的 LOGO**：把图片放进 `static/` 目录，然后：

```python
CORP_LOGO = "my_logo.png"      # 填文件名，不是路径
```

**换成自己的背景**：同样把图片放进 `static/`，然后：

```python
BACKGROUND_IMAGE = "background.jpg"
```

> ⚠️ 图片请使用你有权使用的素材。直接搬运游戏官方美术资源会有版权风险。

#### 10.2 Web 服务

| 变量 | 默认值 | 说明 |
| --- | --- | --- |
| ★ `HOST` | `127.0.0.1` | 监听地址，`0.0.0.0` 表示局域网可访问 |
| ★ `PORT` | `5000` | 监听端口，被占用时改 5001 / 8000 |
| `AUTO_OPEN_BROWSER` | `True` | 启动后是否自动打开浏览器 |
| ★ `SECRET_KEY` | 占位值 | Flask 会话密钥。**公网部署必须改成随机字符串** |

生成随机 SECRET_KEY：

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

#### 10.3 数据目录

| 变量 | 默认值 | 说明 |
| --- | --- | --- |
| `DATA_FOLDER` | `""` | 自定义数据目录，留空用默认 |

`get_data_folder()` 的解析优先级：

```
1. 环境变量 BR_DATA_DIR          ← 最高优先级
2. config.DATA_FOLDER（非空时）
3. 打包成 EXE 运行 → EXE 同级目录 / data
4. 源码运行       → 项目根目录 / data
```

**用法举例**：把数据和报表放到 OneDrive 同步目录，实现自动备份：

```bash
# Windows（命令提示符，临时生效）
set BR_DATA_DIR=D:\OneDrive\br-data
python app.py
```

#### 10.4 数据源接口

| 变量 | 说明 |
| --- | --- |
| `BR_API_BASE` | BR 详情接口，`{}` 会被替换成 BR ID |
| `KM_API_BASE` | 击杀邮件接口，`{}` 会被替换成 killmail ID |
| `ESI_BASE` | EVE 官方 ESI 根地址，供 `ship_id.py` 使用 |
| `BR_LINK_PATTERN` | BR 链接校验正则，**前后端共用** |
| `REQUEST_TIMEOUT` | 单次请求超时秒数，网络慢时调大到 20~30 |

`BR_LINK_PATTERN` 比较特别：它在 Python 端用于 `re.match`，
同时通过 `| tojson` 注入到前端 JS 的 `new RegExp(...)`。
**改一处，前后端同时生效**，不会出现两边规则不一致的问题。

#### 10.5 报告输出

| 变量 | 默认值 | 说明 |
| --- | --- | --- |
| `REPORT_FILE_PREFIX` | `战斗报告` | 报表文件名前缀 |
| `CHART_TOP_N` | `10` | 图表展示前 N 名 |

最终文件名形如：`战斗报告_20250925_测试.xlsx`

#### 10.6 出勤统计

| 变量 | 默认值 | 说明 |
| --- | --- | --- |
| ★ `EVENT_TYPES` | `["反收割", "小队活动", "会战", "集结待命"]` | 出勤页「报告类型」下拉选项 |

改成你自己军团的叫法即可，例如：

```python
EVENT_TYPES = ["反收割", "小队活动", "军团会战", "护卫任务", "集结"]
```

### 11. 需自行修改清单

**定位方式**：在项目根目录执行

```bash
grep -rn "TODO(通用版)\|【需修改】" --include="*.py" --include="*.html" --include="*.txt" .
```

> 行号基于 v1.1 初版，改动后可能偏移，**以搜索到的标记为准**。

#### 11.1 必须修改（影响功能正确性）

| 位置 | 文件 | 行 | 要做什么 |
| --- | --- | --- | --- |
| ① | 网页「军团管理」 | — | 添加你自己军团的 ID 与名称，否则无法统计 |
| ② | `config.py` | 24 | `SITE_NAME` 改成你的系统名 |
| ③ | `config.py` | 30 | `FOOTER_TEXT` 改成你的页脚文案 |
| ④ | `config.py` | 55 | `SECRET_KEY` 改成随机字符串（公网必改） |

#### 11.2 建议修改

| 位置 | 文件 | 行 | 要做什么 |
| --- | --- | --- | --- |
| ⑤ | `config.py` | 34 | `CORP_LOGO` 换自己的 LOGO 或填 `""` 隐藏 |
| ⑥ | `config.py` | 38 | `BACKGROUND_IMAGE` 换自己的背景图 |
| ⑦ | `config.py` | 46 / 49 | `HOST` / `PORT` |
| ⑧ | `config.py` | 125 | `EVENT_TYPES` 活动类型选项 |
| ⑨ | `version_info.txt` | 4 起 | EXE 属性里的公司名 / 版权 |
| ⑩ | `build.py` | 27 / 31 / 34 / 37 | EXE 名称、图标、应用名、版本信息文件 |

#### 11.3 可选 / 按需修改

| 文件 | 行 / 位置 | 说明 |
| --- | --- | --- |
| `config.py` | 64 | `DATA_FOLDER` 自定义数据目录 |
| `config.py` | 94 / 97 / 100 | 数据源接口地址 |
| `config.py` | 103 | `BR_LINK_PATTERN` 链接校验规则 |
| `config.py` | 106 | `REQUEST_TIMEOUT` 超时秒数 |
| `config.py` | 114 | `REPORT_FILE_PREFIX` 文件名前缀 |
| `config.py` | 117 | `CHART_TOP_N` 图表前 N 名 |
| `attendance.py` | 第 25 行 `STORE_FILENAME` | 出勤数据文件名 |
| `templates/base.html` | 13 附近 | CDN 地址（内网环境换成本地文件） |
| `templates/attendance.html` | 24 | 出勤类型下拉框 |
| `tools/make_icon.py` | 22 | 图标配色 |
| `app.py` | 316 | 「未配置军团」的提示文案 |
| `app.py` | 474 | 角色管理扩展点（白名单等） |

#### 11.4 资源替换对照表

| 想换什么 | 操作步骤 |
| --- | --- |
| 导航栏 LOGO | 图片放 `static/` → 改 `config.CORP_LOGO` 为文件名 |
| 页面背景 | 图片放 `static/` → 改 `config.BACKGROUND_IMAGE` |
| 浏览器图标 + EXE 图标 | 改 `tools/make_icon.py` 配色 → `python tools/make_icon.py`；或直接替换 `static/favicon.ico` 与 `app_icon.ico` |
| 出勤活动类型 | 改 `config.EVENT_TYPES` |

---

## 第三部分 · 功能详解

### 12. 战斗统计分析

入口：首页（导航栏「战斗统计分析」）。

#### 12.1 表单字段说明

| 字段 | 必填 | 说明 |
| --- | --- | --- |
| 选择报告日期 | ✅ | 用于生成文件名，格式 `YYYYMMDD` |
| 自定义报告名称 | ✅ | 文件名后缀，最长 50 字，非法字符自动替换为 `_` |
| 粘贴BR链接 | ✅ | 每行一个，格式必须为 `https://br.evetools.org/br/<ID>` |

#### 12.2 操作步骤

1. 选日期、填名称；
2. 粘贴 BR 链接（每行一个）——输入框下方会**实时显示**「共 N 条链接，格式全部正确 / 其中 M 条无效」；
3. 点「生成报告」——进度条会走「初始化 → 分析战斗数据 → 生成最终报告」，随后页面自动刷新；
4. 页面出现绿色提示「BR报告生成完毕」后，点「下载最新报告」。

#### 12.3 统计口径

只统计 `军团管理` 里列出的军团的成员。对每条击杀邮件：

- 你方成员出现在**攻击者**列表 → 记录其舰船、累加其伤害；
- 你方成员是**受害者** → 记录出击次数与舰船；
- 攻击者中 `blow = true`（打出致命一击）→ 该角色击杀数 +1；
- 同一场战斗同一角色只计一次出击。

#### 12.4 生成的文件

见[第 15 章](#15-报表文件说明)。

### 13. 出勤统计分析

#### 13.1 入口地址特别说明

由于 Flask 蓝图前缀叠加（`url_prefix='/attendance'` + 蓝图内 `@route('/attendance')`），
出勤页的**真实地址是**：

```
http://127.0.0.1:5000/attendance/attendance
```

**直接访问 `/attendance` 会返回 404，这是预期行为，不是故障。**
正常从导航栏「出勤统计分析」点进去即可。

#### 13.2 表单字段说明

| 字段 | 必填 | 说明 |
| --- | --- | --- |
| 报告日期 | ✅ | 默认自动填今天 |
| 报告类型 | ✅ | 下拉选项来自 `config.EVENT_TYPES` |
| 舰队说明 | ❌ | 自由文本，例如「护送运输队」 |
| 角色列表 | ✅ | 每行一个角色名 |

#### 13.3 操作步骤

1. 选日期、选类型；
2. 粘贴本次出队的成员名单（每行一个）；
3. 点「添加舰队角色」→ 弹窗提示「已成功添加 N 个角色」，输入框自动清空，光标回到输入框；
4. 右侧表格**实时刷新**出勤次数排行，底部显示合计；
5. 继续提交下一场，重复 1~4；
6. 需要存档时点「导出Excel」。

**去重规则**：同一份名单里重复出现的角色只算一次。
所以「Pilot Alpha / Pilot Beta / Pilot Alpha」提交后，会显示添加了 2 个角色。

#### 13.4 数据持久化

出勤数据保存在 `data/attendance.json`，**关掉程序再打开数据仍在**。
页面加载时会自动检测是否有历史数据，有则直接开放导出按钮并刷新表格。

### 14. 管理后台

导航栏 → **管理**（下拉菜单）。

#### 14.1 军团管理

- **作用**：维护「要统计哪些军团」
- **文件**：`data/corp_id.txt`
- **提示**：通用版默认为空，首次使用必须添加

字段：军团ID（纯数字）、军团名称（仅展示用）。

已添加的军团以列表展示，带条目计数，可逐个删除（会二次确认）。

#### 14.2 角色管理

- **作用**：补充/修正角色 ID 与名称映射
- **文件**：`data/char_id.txt`

> **通常不需要手动维护**。战斗统计过程中，程序会自动把 BR 里出现的本军团成员
> 补进 `char_id.txt`。
>
> 只有当报表里出现 `Unknown(数字ID)`，或你想给某个未出现过的人预置名称时，
> 才需要手工添加一条。

#### 14.3 历史报告

- **作用**：查看、下载、删除已生成的报表
- **位置**：`data/reports/`
- 按生成时间倒序排列，删除时会二次确认

#### 14.4 系统日志

- **作用**：排查报错
- **文件**：`data/system.log`
- 记录每一条 `SUCCESS` / `ERROR` / `INFO`，含时间戳

日志示例：

```
[2025-09-25 10:12:31.220000] SUCCESS: Report generated: 战斗报告_20250925_测试.xlsx
[2025-09-25 10:15:02.100000] ERROR: Invalid BR link: https://br.evetools.org/br/xxx
```

另有一份更详细的运行日志 `data/detailed.log`（由 Python logging 写入）。

### 15. 报表文件说明

#### 15.1 战斗统计报表

文件名：`战斗报告_<日期>_<自定义名称>.xlsx`
默认位置：`data/reports/`

**工作表「战斗统计」**

| 列 | 字段 | 说明 |
| --- | --- | --- |
| A | 排名 | 按总伤害降序 |
| B | 角色 | 角色名 |
| C | 出击次数 | 出现在战斗记录中的次数 |
| D | 击杀数 | 致命一击次数 |
| E | 总伤害 | 累计伤害 |
| F | 使用舰船 | 去重后的舰船列表，顿号分隔 |

**附带图表**

- `H2` 位置：「总伤害排名前10」柱状图
- `H18` 位置：「击杀数排名前10」柱状图

图表前 N 名由 `config.CHART_TOP_N` 控制。

#### 15.2 出勤统计报表

文件名：`舰队出勤统计_<时间戳>.xlsx`
下载位置：浏览器默认下载目录

| 工作表 | 内容 |
| --- | --- |
| 出勤统计 | 角色名、出现次数（降序） |
| 历史记录 | 提交时间、活动日期、活动类型、角色数量、活动说明、角色列表 |

---

## 第四部分 · 打包与发布

### 16. 打包为 EXE

适合把工具发给不会 Python 的队友。

```bash
# 1) 装打包依赖
pip install pyinstaller pillow

# 2)（可选）重新生成图标
python tools/make_icon.py

# 3) 打包
python build.py
```

产物：`dist/BRStatisticsTool.exe`

打包配置在 `build.py` 顶部：

| 常量 | 说明 |
| --- | --- |
| `EXE_NAME` | EXE 文件名 |
| `ICON_FILE` | 图标文件名（`.ico`） |
| `APP_NAME` | Windows 属性里的应用名（默认取 `config.SITE_NAME`） |
| `VERSION_FILE` | 版本信息文件（默认 `version_info.txt`） |

打包脚本会自动：

1. 清理 `build/`、`dist/` 和旧 `.spec`；
2. 把 `templates/`、`static/`、`data/` 一并打进 EXE；
3. 声明 `jinja2 / werkzeug / flask / openpyxl / pandas` 等隐藏导入；
4. 应用图标与版本信息。

**首次运行 EXE**：Windows 可能提示「已保护你的电脑」→ 点「更多信息 → 仍要运行」。

> 💡 打包后**数据写在 EXE 同级目录**的 `data/` 下，不在临时目录，
> 所以关掉程序数据不会丢。这也是通用版修复的问题之一。

### 17. 分发与迁移

发给别人时，拷贝：

```
BRStatisticsTool.exe
data/                    # 可只带 ship_id.txt
```

**迁移到新电脑**：把整个 `data/` 目录拷到新机器的同级位置即可，
历史报表和出勤记录都会一起过去。

### 18. 上传到 GitHub

项目已内置 `.gitignore`，会自动排除：

- `build/`、`dist/`、`*.spec`（构建产物）
- `data/corp_id.txt`、`data/char_id.txt`、`data/br_links.txt`（用户数据）
- `data/attendance.json`、`data/*.log`、`data/reports/*.xlsx`
- `.venv/`、`.env`、`*.pem`、`*.key`

标准流程：

```bash
git init
git add .
git commit -m "feat: 通用版 BR 战斗统计系统"
git branch -M main
git remote add origin https://github.com/<用户名>/<仓库名>.git
git push -u origin main
```

#### ⚠️ Token 安全

**不要把 Token 写进代码、脚本、`.git/config` 或提交进仓库。**

如果 Token 曾出现在聊天记录、日志或截图里，请立刻：

1. 打开 GitHub `Settings → Developer settings → Personal access tokens`
2. 找到对应 Token → **Revoke（吊销）**
3. 重新生成一个，权限只勾选必需的（`Contents: Read and write`）

推送完成后检查 `.git/config` 是否残留 Token：

```bash
grep -c "github_pat\|ghp_" .git/config || echo "clean"
```

#### 推送失败怎么办

若 `git push` 报 `schannel: server closed abruptly`、`SSL_ERROR_SYSCALL`
或长时间挂起，通常是 **`github.com` 的 git 端口被网络阻断**（但 `api.github.com` 可访问）。
此时可改用 GitHub REST API 推送，详见项目内置技能 `github-push-via-api`。

---

## 第五部分 · 运维与排错

### 19. 数据备份与恢复

#### 需要备份什么

| 文件 | 重要程度 | 说明 |
| --- | --- | --- |
| `data/attendance.json` | 🔴 高 | 出勤记录，丢了要重新录 |
| `data/corp_id.txt` | 🔴 高 | 军团配置 |
| `data/char_id.txt` | 🟡 中 | 角色映射，可自动重建 |
| `data/reports/*.xlsx` | 🟡 中 | 历史报表 |
| `data/ship_id.txt` | 🟢 低 | 公共数据，可重新拉取 |

**最简单的备份方式**：整个 `data/` 目录复制一份。

**自动备份**：用环境变量把数据目录指到网盘同步目录：

```bash
set BR_DATA_DIR=D:\OneDrive\br-data
python app.py
```

#### 恢复

把备份的 `data/` 覆盖回去即可，无需其他操作。

### 20. 升级流程

1. 备份 `data/`（见上一章）；
2. 用新版代码覆盖除 `data/` 以外的文件；
3. **如果新版改过 `config.py`**，需要把你自己改过的配置项重新填一遍；
4. 重新 `pip install -r requirements.txt`（依赖可能有变）；
5. 启动，跑一次 `python tools/smoke_test.py` 确认正常。

### 21. 错误信息速查表

| 页面提示 / 控制台信息 | 原因 | 解决办法 |
| --- | --- | --- |
| `Invalid links: xxx` | BR 链接格式不合规 | 检查格式必须是 `https://br.evetools.org/br/<ID>`，不能带多余参数 |
| `Invalid BR link: xxx` | 链接格式对但接口取不到 | BR 已过期 / 接口临时不可用，刷新重试 |
| `Failed to access: xxx` | 同上 | 同上 |
| `尚未配置任何军团 ID，请先到「军团管理」页面添加…` | `corp_id.txt` 为空 | 到「军团管理」添加军团 |
| `Date and name are required` | 日期或名称没填 | 两个都是必填 |
| `Please enter valid BR links` | 链接框为空 | 填入至少一条链接 |
| `Report generation failed: xxx` | 生成过程异常 | 看「系统日志」里的详细信息 |
| `[提示] 数据文件 corp_id.txt 不存在，已按空数据处理` | 首次运行，文件还没生成 | **正常提示**，配置军团后自动创建 |
| 报表里 `Unknown(2118200607)` | 角色映射缺失 | 下次统计会自动补全；也可到「角色管理」手动加 |
| 报表里 `Unknown(12345)`（舰船） | 有新船未收录 | 执行 `python ship_id.py` |
| 请求一直超时 | 网络/代理问题 | 关 VPN；或调大 `REQUEST_TIMEOUT` |
| 页面 404（访问 `/attendance`） | 蓝图前缀叠加 | 用 `/attendance/attendance` 或从导航栏进入 |
| 端口被占用 | 5000 被别的程序占了 | 改 `config.PORT` |

### 22. 常见问题 FAQ

**Q1：点「出勤统计分析」后地址栏是 `/attendance`，页面 404？**

蓝图前缀叠加导致。真实地址是 `/attendance/attendance`。
正常从导航栏点进去不会有问题，详见[第 13.1 节](#131-入口地址特别说明)。

**Q2：为什么第一次生成报告就报错说没配置军团？**

通用版默认军团列表为空。这是刻意设计——避免生成一份全空白报表让你以为是程序坏了。
去「管理 → 军团管理」加一条就行。

**Q3：能不能一次统计多个军团？**

可以。在「军团管理」里把相关军团都加进去，统计时会一起收录。
统计范围是所有已配置军团的并集。

**Q4：BR 链接要一个个粘吗？可以批量吗？**

直接多行粘贴即可，每行一个。程序会依次处理，验证阶段发现有问题的链接会直接指出是哪一条。

**Q5：为什么我的伤害数字和游戏里不一样？**

本系统统计的是 **BR 记录里的伤害**，不是游戏内的伤害统计。
口径差异来自：BR 只覆盖你提供的这几个链接；致命一击才计击杀数；
同一场战斗同一角色只计一次出击。

**Q6：出勤数据会丢吗？**

不会。存在 `data/attendance.json`，重启不丢。备份这个文件即可。

**Q7：局域网内其他人怎么访问？**

`HOST` 改成 `"0.0.0.0"`，别人访问 `http://<你的内网IP>:5000`。
需要在防火墙放行端口。公网暴露前**务必修改 `SECRET_KEY`**。

**Q8：能部署到公网吗？**

技术上可以，但本程序**没有登录认证**，公网暴露等于任何人都能用。
如确需公网使用，请先加认证（见[第 24.4 节](#244-想加登录认证怎么做)）并配 HTTPS。

**Q9：报表里舰船显示成数字怎么办？**

说明有新船没收录。执行 `python ship_id.py` 重新拉取舰船数据，
生成后覆盖 `data/ship_id.txt`（脚本默认已直接写入该位置）。

**Q10：能换掉页面上的背景图和 LOGO 吗？**

可以。图片放 `static/`，然后改 `config.CORP_LOGO` / `config.BACKGROUND_IMAGE`。
注意使用你有权使用的素材。

**Q11：程序支持 macOS / Linux 吗？**

支持运行。但 `build.py` 打包 EXE 只对 Windows 有效；
在 Linux/macOS 上打包得到的是对应平台的可执行文件。

**Q12：想让程序常驻后台，不要每次开命令行窗口？**

Windows 下可以用 `pythonw app.py`（无控制台窗口），
或者打包成 EXE 后加到「启动」文件夹。注意这样跑起来后就没有日志窗口了，
排查问题时还是用命令行方式启动更方便。

---

## 第六部分 · 开发者指引

### 23. 代码结构与关键函数

#### 23.1 app.py

| 函数 / 路由 | 作用 |
| --- | --- |
| `inject_globals()` | 模板全局变量注入（品牌信息、导航项） |
| `data_path(*parts)` | 拼接数据目录下的路径（**统一用它，别写 `__file__`**） |
| `load_ids(filename)` | 读取 `ID_名称` 映射；跳过空行与 `#` 注释 |
| `save_ids(filename, data)` | 写回映射文件 |
| `resolve_report_path(filename)` | **安全校验**：只允许 `data/reports` 下真实存在的 `.xlsx` |
| `process_killmail(...)` | 解析单条击杀邮件，累加数据 |
| `generate_final_data(...)` | 汇总成报表结构 |
| `/` | 首页：提交 BR 链接、生成报告 |
| `/download-report/<filename>` | 下载报告（带白名单校验） |
| `/delete-report/<filename>` | 删除报告（带白名单校验） |
| `/manage-chars` / `/manage-corps` | 角色 / 军团管理 |
| `/history` / `/logs` / `/docs` | 历史 / 日志 / 文档页 |

#### 23.2 attendance.py

| 函数 | 作用 |
| --- | --- |
| `_store_path()` | 出勤数据文件绝对路径 |
| `_load_submissions()` | 读取全部出勤记录（容错：文件损坏返回空列表） |
| `_save_submissions()` | 写回全部出勤记录 |
| `_build_stats()` | 按角色名汇总次数并排序 |
| `/submit_fleet` | 提交名单（**提交时去重**） |
| `/get_stats` / `/get_history` | 统计数据 / 历史记录 |
| `/export_stats` | 导出 Excel（`BytesIO` 内存流，不落盘） |

#### 23.3 安全相关的三处实现（**改动时不要退化**）

1. **路径穿越防护** — `resolve_report_path()`
   只接受「不含路径分隔符 + 以 `.xlsx` 结尾 + 在 `data/reports` 目录列表里」的文件名。
   > ⚠️ **不要改成 `werkzeug.secure_filename`**——它会把中文报表名清空，导致下载失效。

2. **XSS 防护** — `templates/attendance.html` 的 `escapeHtml()`
   角色名来自用户输入，写入 DOM 前必须转义。

3. **前端校验规则同源** — `config.BR_LINK_PATTERN` 通过 `| tojson` 注入 JS，
   避免前后端规则不一致。

### 24. 二次开发指引

#### 24.1 加一个新的配置项

1. 在 `config.py` 对应章节添加变量，注释标 `【需修改】` 或说明用途；
2. 若模板需要，在 `app.py` 的 `inject_globals()` 里加一行；
3. 更新 `操作手册.md` 第 10 / 11 章与 `OPERATION_MANUAL_EN.md` 对应章节。

#### 24.2 加一个新页面

```python
@app.route('/my-page')
def my_page():
    return render_template('my_page.html', ...)
```

模板继承 `base.html` 即可自动获得导航、页脚、样式：

```jinja
{% extends "base.html" %}
{% block content %}
  ...你的内容...
{% endblock %}
```

导航项在 `app.py` 的 `inject_globals()` 里加。

#### 24.3 统计口径要改（比如「击杀数」改成统计参与击杀）

修改 `process_killmail()` 中致命一击那一段：

```python
# 现状：只有 blow=true 才 +1
for attacker in kb_data.get('atts', []):
    if attacker.get('blow', True):
        ...
        char_dir[char_str][0][0] += 1
        break
```

改成统计「参与击杀」（去掉 blow 判断、去掉 break）即可，但要自己评估数据口径是否合理。

#### 24.4 想加登录认证怎么做

程序原本预留了登录相关的注释但未实现。推荐做法：

1. `pip install flask-login`；
2. 在 `app.py` 加 `LoginManager`；
3. 用一个简单的用户表（可以就放在 `data/` 下的 JSON 文件里）；
4. 给所有路由加 `@login_required`。

**不要**用「把用户名密码写在代码里」的方式——这是最容易被扫到的漏洞。

#### 24.5 改完必须验证

```bash
python -m py_compile app.py config.py attendance.py    # 语法检查
python tools/smoke_test.py                              # 25 项冒烟测试
```

`smoke_test.py` 会把数据目录临时指到系统临时目录，
**不会污染你的真实 `data/`**，可以放心反复运行。

### 25. 通用版变更记录

#### 25.1 相比原定制版清除的内容

| 项目 | 处理方式 |
| --- | --- |
| 军团头像 `static/corp_logo.png` | 删除，换成中性的 `site_logo.svg` |
| 军团专属图标 `cosmic_wanderers_icon.ico` | 删除，换成 `app_icon.ico` / `favicon.ico` |
| 军团背景图（2.3MB 官方美术素材） | 删除，改为 CSS 深空渐变 |
| 军团名 `Cosmic-Wanderers` | 全部替换为 `config.SITE_NAME` |
| 军团 ID | `data/corp_id.txt` 清空 |
| 260 条真实角色 ID / 名称 | `data/char_id.txt` 清空 |
| 3 条真实 BR 链接 | `data/br_links.txt` 清空 |
| 战斗与报表日志 | `system.log`、`detailed.log` 删除 |
| 作者署名 | 已移除 |
| 遗留产物 `output.xlsx`、`__pycache__/` | 删除 |

#### 25.2 修复的问题

| # | 问题 | 修复 |
| --- | --- | --- |
| 1 | 出勤数据只存内存，重启即丢 | 持久化到 `data/attendance.json` |
| 2 | 导出 Excel 在 CWD 留 `output.xlsx` 垃圾 | 改 `BytesIO` 内存流 |
| 3 | 报告接口存在目录穿越漏洞 | 增加白名单校验 |
| 4 | 出勤表格 `innerHTML` 拼角色名，XSS 风险 | 增加 `escapeHtml()` |
| 5 | 打包后数据写进临时目录，重启丢失 | 数据目录改为 EXE 同级，支持 `BR_DATA_DIR` |
| 6 | 文件读写未指定编码 | 统一 `encoding='utf-8'` |
| 7 | 前后端各写一份 BR 校验正则 | 统一由 `config.BR_LINK_PATTERN` 注入 |
| 8 | 四个管理页面没有导航入口 | 导航栏新增「管理」下拉 |
| 9 | 表单出错后已填内容丢失 | 回显日期/名称/链接，并保存上次链接 |
| 10 | Bootstrap 引入双版本、Font Awesome 重复 | 清理为单一版本，统一用 bootcdn |
| 11 | 未配置军团时生成空白报表 | 改为明确报错并引导配置 |
| 12 | 出勤名单重复角色重复计数 | 提交时按顺序去重 |
| 13 | 删除按钮 `onclick` 拼文件名，含引号报错 | 改 `data-*` 属性 + `encodeURIComponent` |
| 14 | 空列表页面一片空白 | 增加空状态提示与条目计数 |
| 15 | 残留个人绝对路径 | 改为基于 `config.get_data_folder()` |
| 16 | `docs.html` 标签嵌套错误 | 重写为规范结构 |

#### 25.3 新增内容

- `config.py` — 统一配置入口
- `README.md` — 项目简介（中文）
- `操作手册.md` — 中文详细操作手册（本文件）
- `OPERATION_MANUAL_EN.md` — 英文详细操作手册
- `data/README.md` — 数据文件格式说明
- `tools/make_icon.py` — 图标生成脚本
- `tools/smoke_test.py` — 冒烟测试脚本（25 项断言）
- `.gitignore`、`requirements.txt`
- 中性的 LOGO / favicon / EXE 图标

---

## 附录

### 附录 A：数据文件格式

统一约定：**UTF-8 编码，每行一条记录，以第一个下划线 `_` 分隔 ID 与名称。**
空行和以 `#` 开头的注释行会被忽略。

#### A.1 `corp_id.txt`

```
# 军团ID_军团名称
98769610_Example Corporation
12345678_Another Corporation
```

#### A.2 `char_id.txt`

```
2118200607_Character Name
2113436255_Another Pilot
```

#### A.3 `ship_id.txt`

```
582_Bantam
583_Condor
584_Griffin
```

公共数据，已预置 6000+ 条。有新船时执行 `python ship_id.py` 更新。

#### A.4 `br_links.txt`

```
https://br.evetools.org/br/68d956f22213cf0012ace2af
https://br.evetools.org/br/68d6bab14ff6700012d935c5
```

用于页面回显上次提交的链接。

#### A.5 `attendance.json`

```json
[
  {
    "timestamp": "2025-09-25T10:12:31.220000",
    "date": "2025-09-25",
    "event_type": "小队活动",
    "note": "护送运输队",
    "members": ["Pilot Alpha", "Pilot Beta"],
    "member_count": 2
  }
]
```

> 请勿手工损坏该文件，否则出勤数据会被当作空数据。

### 附录 B：配置项速查表

```python
# ---- 品牌 ----
SITE_NAME         = "EVE Online BR 战斗统计系统"
SITE_VERSION      = "1.1"
FOOTER_TEXT       = "EVE Online BR 战斗统计系统"
CORP_LOGO         = "site_logo.svg"        # ""=不显示
BACKGROUND_IMAGE  = ""                     # ""=用内置渐变

# ---- 服务 ----
HOST              = "127.0.0.1"            # "0.0.0.0"=局域网可访问
PORT              = 5000
AUTO_OPEN_BROWSER = True
SECRET_KEY        = "please-change-this-secret-key"   # 公网必改

# ---- 数据目录 ----
DATA_FOLDER       = ""                     # ""=默认；环境变量 BR_DATA_DIR 优先级更高

# ---- 接口 ----
BR_API_BASE       = "https://br.evetools.org/api/v1/composition/get/{}"
KM_API_BASE       = "https://kb.evetools.org/api/v1/killmails/{}"
ESI_BASE          = "https://esi.evetech.net/latest"
BR_LINK_PATTERN   = r"^https://br\.evetools\.org/br/[\w-]+$"
REQUEST_TIMEOUT   = 10

# ---- 报表 ----
REPORT_FILE_PREFIX = "战斗报告"
CHART_TOP_N        = 10

# ---- 出勤 ----
EVENT_TYPES = ["反收割", "小队活动", "会战", "集结待命"]
```

### 附录 C：术语表

| 术语 | 英文 | 含义 |
| --- | --- | --- |
| BR | Battle Report | 战斗报告，记录一场战斗参与方与结果的链接 |
| KM | Killmail | 击杀邮件，EVE 中每次击杀都会生成一条记录 |
| 军团 | Corporation | EVE 中的基础组织单位 |
| 联盟 | Alliance | 由多个军团组成的更大组织 |
| 出击次数 | Sorties | 该角色出现在战斗记录中的次数 |
| 致命一击 | Final Blow | 打出最后一击，BR 中标记为 `blow = true` |
| 舰船 | Ship | EVE 中的可驾驶载具 |
| ESI | EVE Swagger Interface | EVE 官方开放 API |
| 蓝图 | Blueprint | Flask 中用于组织路由的组件 |

---

### 附：一分钟上手流程

```
1. pip install -r requirements.txt
2. python app.py
3. 浏览器 → 管理 → 军团管理 → 添加你的「军团ID_军团名称」
4. 首页 → 填日期、名称 → 粘贴 BR 链接 → 生成报告 → 下载
5. 出勤统计 → 导航栏「出勤统计分析」→ 录名单 → 导出Excel
```

---

> 本文档与 [OPERATION_MANUAL_EN.md](#english-operation-manual) 内容一一对应。
> 如果两份文档出现不一致，**以中文本为准**（中文版是原始维护版本）。

---

## English Operation Manual

## EVE Online BR Combat Statistics System · Detailed Operation Manual

> **English version** ｜ 中文版 → [操作手册.md](#中文操作手册)
> Applies to: v1.1 Generic Edition
> Intended for: deployers (first-time users), daily operators, and developers extending the tool

---

### How to read this manual

| Who you are | Read this |
| --- | --- |
| First time here, just want to get it running | Part 1 → Part 2 Ch. 6–9 → Part 3 |
| Need to rebrand, reconfigure, change ports | Part 2 Ch. 10–11 |
| Want to build an EXE and share it | Part 4 |
| Something broke | Part 5 Ch. 21–22 |
| Want to modify the code | Part 6 |

Every `[MUST EDIT]` marker below corresponds to a real `TODO(通用版)` / `【需修改】`
comment in the source. Search for those strings to jump straight to the code.

---

### Table of Contents

**Part 1 · Understanding the System**
- [1. What this system is](#1-what-this-system-is)
- [2. Feature list](#2-feature-list)
- [3. How it works — data flow](#3-how-it-works--data-flow)
- [4. Directory structure](#4-directory-structure)

**Part 2 · Deployment and Configuration**
- [5. Requirements](#5-requirements)
- [6. Installation (from source)](#6-installation-from-source)
- [7. Starting and stopping](#7-starting-and-stopping)
- [8. Three deployment modes](#8-three-deployment-modes)
- [9. Mandatory first-time setup](#9-mandatory-first-time-setup)
- [10. Complete configuration reference (config.py)](#10-complete-configuration-reference-configpy)
- [11. What you must edit yourself](#11-what-you-must-edit-yourself)

**Part 3 · Feature Walkthroughs**
- [12. Combat statistics](#12-combat-statistics)
- [13. Attendance statistics](#13-attendance-statistics)
- [14. Admin pages](#14-admin-pages)
- [15. Report files explained](#15-report-files-explained)

**Part 4 · Packaging and Publishing**
- [16. Building an EXE](#16-building-an-exe)
- [17. Distribution and migration](#17-distribution-and-migration)
- [18. Publishing to GitHub](#18-publishing-to-github)

**Part 5 · Operations and Troubleshooting**
- [19. Backup and restore](#19-backup-and-restore)
- [20. Upgrade procedure](#20-upgrade-procedure)
- [21. Error message reference](#21-error-message-reference)
- [22. FAQ](#22-faq)

**Part 6 · Developer Guide**
- [23. Code structure and key functions](#23-code-structure-and-key-functions)
- [24. Extending the system](#24-extending-the-system)
- [25. Generic edition changelog](#25-generic-edition-changelog)

**Appendices**
- [A. Data file formats](#appendix-a-data-file-formats)
- [B. Configuration cheat sheet](#appendix-b-configuration-cheat-sheet)
- [C. Glossary](#appendix-c-glossary)

---

## Part 1 · Understanding the System

### 1. What this system is

**EVE Online BR Combat Statistics System** is a local web tool that runs on your own computer.

Its core job: **turn a list of BR (Battle Report) links into a damage leaderboard.**

Combat records in EVE Online live on third-party sites, and a single large engagement can
generate dozens or hundreds of BR links. Manually working out who contributed and who
did nothing is simply not practical. This tool automates that:

```
You paste a batch of BR links  →  the app fetches each combat record  →  keeps only members of the corporations you specified  →  produces an Excel report
```

**It is not**: an official EVE tool, a game modification, an account reader, or anything that
requires you to log in. It does exactly one thing — **aggregate public data**.

### 2. Feature list

| # | Feature | Where | Output |
| --- | --- | --- | --- |
| 1 | Combat statistics | Home page → "战斗统计分析" | Excel with charts (rank / pilot / sorties / kills / damage / ships) |
| 2 | Attendance statistics | Navbar → "出勤统计分析" | Attendance ranking + Excel (with a history sheet) |
| 3 | Corporation management | Manage → Corporation | Maintains *which corporations to track* |
| 4 | Character management | Manage → Character | Maintains character ID ↔ name mapping |
| 5 | Report history | Manage → History | View / download / delete past reports |
| 6 | System log | Manage → System log | Troubleshooting source |

> The web UI ships in Chinese. English-speaking users can still use every feature —
> this chapter names each menu so you can find it. If you want an English UI,
> see [24.6](#246-translating-the-ui-into-another-language).

### 3. How it works — data flow

#### 3.1 Combat statistics pipeline

```
 ① You paste BR links
        │
        ▼
 ② Format validation (regex, checked on both the client and the server)
        │  invalid → red inline message, no requests fired
        ▼
 ③ GET https://br.evetools.org/api/v1/composition/get/{BR_ID}
        │  returns this fight's relateds → kms[] (list of killmails)
        ▼
 ④ For each killmail: GET https://kb.evetools.org/api/v1/killmails/{KILL_ID}
        │
        ├── Extract atts[] (attackers) and victim
        ├── Compare each corp (corporation ID) against data/corp_id.txt
        │      ├─ match    → accumulate sorties / damage for that pilot
        │      └─ no match → skip entirely (this is how "only our own" works)
        ├── If an attacker has blow (final blow) = true → that pilot's kills +1
        └── Unknown pilot encountered → auto-appended to data/char_id.txt
        ▼
 ⑤ generate_final_data() aggregates: sorties / kills / total damage / ships (deduplicated)
        ▼
 ⑥ openpyxl writes the workbook: data sheet + Damage Top10 chart + Kills Top10 chart
        ▼
 ⑦ Saved to data/reports/, and a "download latest report" button appears
```

**Key metrics**

| Metric | Meaning |
| --- | --- |
| Sorties | Number of times the pilot appears across the BRs you submitted (once per fight) |
| Kills | Number of **final blows**, not number of fights participated in |
| Total damage | Sum of damage dealt across all killmails |
| Ships used | Deduplicated list of ship names seen |

> ⚠️ **Important**: the scope is determined entirely by `data/corp_id.txt`.
> If it is empty, the app refuses and points you to the configuration page.
> This is **intentional** — it prevents generating a silently blank report that makes
> you think the tool is broken.

#### 3.2 Attendance statistics pipeline

```
 ① Enter date + event type + member list (one name per line)
        │
        ▼
 ② Front-end fetch POST → /attendance/submit_fleet
        │
        ├── Deduplicate in order (a name repeated within one list counts once)
        └── Append to data/attendance.json
        ▼
 ③ Front-end immediately GETs /attendance/get_stats to refresh the live table
        ▼
 ④ Click "导出Excel" → /attendance/export_stats
        └── Workbook built in memory and streamed to the browser — no temp file on disk
```

> Attendance data is **persisted to disk**. Close the app, reopen it — your data is still there.

#### 3.3 Third-party endpoints used

| Purpose | URL | Notes |
| --- | --- | --- |
| BR detail | `https://br.evetools.org/api/v1/composition/get/{id}` | Killmail list for a fight |
| Killmail detail | `https://kb.evetools.org/api/v1/killmails/{id}` | Attacker / victim breakdown |
| Ship data | `https://esi.evetech.net/latest` | Official ESI; used only by `ship_id.py` |

All of these live in `config.py` and can be swapped if an endpoint changes.

### 4. Directory structure

```
br-web/
│
├── app.py                      # Main app: routes + combat logic + Excel generation
├── config.py                   # ★ Single configuration entry point
├── attendance.py               # Attendance blueprint (with JSON persistence)
├── ship_id.py                  # Ship ID → name updater
├── build.py                    # PyInstaller build script
├── version_info.txt            # EXE version resource (company / copyright)
├── requirements.txt            # Dependency list
│
├── 操作手册.md                 # ★ Detailed manual (Chinese)
├── OPERATION_MANUAL_EN.md      # ★ This file (English)
├── README.md                   # Project overview
├── app_icon.ico                # EXE icon
│
├── tools/
│   ├── make_icon.py            # Generates favicon / EXE icon / PNG logo
│   └── smoke_test.py           # Smoke test (25 assertions, isolated data dir)
│
├── data/                       # Runtime data (see data/README.md)
│   ├── ship_id.txt             # Ship mapping (public data, 6000+ entries preloaded)
│   ├── corp_id.txt             # Corporation config ← you fill this in first
│   ├── char_id.txt             # Character mapping (written automatically)
│   ├── br_links.txt            # Last submitted BR links (saved automatically)
│   ├── attendance.json         # Attendance records (created automatically)
│   ├── reports/                # Generated Excel reports
│   └── README.md
│
├── static/
│   ├── site_logo.svg           # Navbar logo (neutral placeholder, replaceable)
│   ├── site_logo.png           # Same, PNG variant
│   ├── favicon.ico             # Browser tab icon
│   └── style.css
│
└── templates/                  # Jinja2 templates
    ├── base.html               # Layout: navbar, footer, background, styles
    ├── index.html              # Combat statistics home page
    ├── attendance.html         # Attendance page
    ├── docs.html               # Built-in help page
    ├── manage_corps.html       # Corporation management
    ├── manage_chars.html       # Character management
    ├── history.html            # Report history
    └── logs.html               # System log
```

---

## Part 2 · Deployment and Configuration

### 5. Requirements

| Item | Minimum | Recommended |
| --- | --- | --- |
| OS | Windows 10 / 11, Linux, macOS | Windows 11 |
| Python | 3.9 | 3.10 – 3.12 |
| RAM | 512 MB | 2 GB |
| Disk | 200 MB (incl. ship data) | 1 GB (room for reports) |
| Network | Reachable `br.evetools.org`, `kb.evetools.org` | Direct, **VPN/proxy off** |
| Browser | Chrome / Edge / Firefox | Chrome / Edge |

**About VPNs**: this matters. The BR endpoints rate-limit frequent requests, and going through
a proxy makes throttling or suspicious-traffic blocks much more likely. The upstream site
itself advises against using a VPN for BR statistics.

### 6. Installation (from source)

```bash
# 1) Enter the project directory
cd br-web

# 2) Create a virtual environment (recommended)
python -m venv .venv

# 3) Activate it
.venv\Scripts\activate           # Windows
# source .venv/bin/activate      # Linux / macOS

# 4) Install dependencies
pip install -r requirements.txt

# 5) Run
python app.py
```

If pip is slow, use a mirror (mainland China):

```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

Dependencies:

| Package | Purpose | Required |
| --- | --- | --- |
| Flask | Web framework | ✅ |
| requests | Calls to third-party endpoints | ✅ |
| openpyxl | Excel + chart generation | ✅ |
| pandas | Attendance Excel export | ✅ |
| pyinstaller | Building the EXE | Build only |
| pillow | Generating icons | Icon work only |

### 7. Starting and stopping

```bash
python app.py
```

- The browser **opens automatically** (controlled by `config.AUTO_OPEN_BROWSER`);
- The console prints the data directory and the URL — **do not close this window**;
- To stop: press `Ctrl + C` in that window, or just close it.

**If the port is taken**, change `PORT` in `config.py` (5001 and 8000 both work).
To find what is holding 5000:

```bash
netstat -ano | findstr :5000     # Windows
lsof -i :5000                    # Linux / macOS
```

### 8. Three deployment modes

#### Mode 1 — Local only (default, simplest)

```python
# config.py
HOST = "127.0.0.1"
PORT = 5000
```

Only you can reach it. No security considerations.

#### Mode 2 — Shared on a LAN

```python
# config.py
HOST = "0.0.0.0"           # listen on all interfaces
PORT = 5000
SECRET_KEY = "<your own random string>"
```

Others browse to `http://<your-LAN-IP>:5000`.
Find your LAN IP with `ipconfig` on Windows (look for "IPv4 Address").
You also need to **allow the port through Windows Firewall** — usually just clicking
"Allow access" the first time Windows prompts about Python.

#### Mode 3 — Public server (careful)

Only if you really need it, and you must:

1. Change `SECRET_KEY` to a random string;
2. Put it behind Nginx with HTTPS;
3. Understand that **this app has no authentication** — exposing it publicly means anyone
   can generate reports and read your history. If you need auth, see
   [24.4](#244-how-do-i-add-login-authentication).

### 9. Mandatory first-time setup

Once installed, **one step remains**: tell the app which corporation to track.

#### 9.1 Configure via the web UI (recommended)

1. Open <http://127.0.0.1:5000>
2. Navbar → **管理 (Manage) → 军团管理 (Corporation)**
3. Fill in:
   - **军团ID (Corporation ID)** — digits only, e.g. `98769610`
   - **军团名称 (Corporation name)** — display only, anything you like
4. Click **添加军团 (Add)**

#### 9.2 How to find your corporation ID

Open <https://zkillboard.com> → search your corporation → open its page and look at the URL:

```
https://zkillboard.com/corporation/98769610/
                                └────┬───┘
                              This is the corporation ID
```

You can also read it off the corporation page on <https://br.evetools.org>.

#### 9.3 Editing the file directly (optional)

Edit `data/corp_id.txt`, one `ID_name` per line:

```
98769610_Example Corporation
12345678_Another Corporation
```

Blank lines and lines starting with `#` are ignored.

#### 9.4 Verifying it worked

Go back to the home page and submit any BR link. If no corporation is configured, the page
shows a red message — *"尚未配置任何军团 ID，请先到「军团管理」页面添加你要统计的军团"*
("No corporation ID configured yet — add the corporation you want to track on the
Corporation page first"). Seeing that message means `corp_id.txt` is still empty.

### 10. Complete configuration reference (config.py)

Everything configurable lives in `config.py`. **Items marked ★ are the ones you will
normally need to change.**

#### 10.1 Branding

| Variable | Default | Description |
| --- | --- | --- |
| ★ `SITE_NAME` | `EVE Online BR 战斗统计系统` | Used for the browser title, navbar, and footer |
| ★ `SITE_VERSION` | `1.1` | Rendered as "<name> v1.1" |
| ★ `FOOTER_TEXT` | same as `SITE_NAME` | Footer text; final output is "<text> ©<year>" |
| ★ `CORP_LOGO` | `site_logo.svg` | Navbar logo. Set to `""` to hide the logo entirely |
| `BACKGROUND_IMAGE` | `""` | Page background. Empty string = built-in deep-space gradient |

**Using your own logo**: drop the image into `static/`, then:

```python
CORP_LOGO = "my_logo.png"      # file name only, not a path
```

**Using your own background**:

```python
BACKGROUND_IMAGE = "background.jpg"
```

> ⚠️ Use artwork you actually have the rights to. Copying official game art carries
> copyright risk.

#### 10.2 Web service

| Variable | Default | Description |
| --- | --- | --- |
| ★ `HOST` | `127.0.0.1` | Bind address; `0.0.0.0` makes it LAN-reachable |
| ★ `PORT` | `5000` | Port; change if taken |
| `AUTO_OPEN_BROWSER` | `True` | Open the browser on startup |
| ★ `SECRET_KEY` | placeholder | Flask session key. **Must be changed for public deployments** |

Generate a random key:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

#### 10.3 Data directory

| Variable | Default | Description |
| --- | --- | --- |
| `DATA_FOLDER` | `""` | Custom data directory; empty = default |

Resolution order in `get_data_folder()`:

```
1. Environment variable BR_DATA_DIR        ← highest priority
2. config.DATA_FOLDER (when non-empty)
3. Frozen EXE       → <dir of exe>/data
4. Running from source → <project root>/data
```

**Example**: keep data and reports inside a OneDrive folder for automatic backup:

```bash
# Windows (cmd, session-scoped)
set BR_DATA_DIR=D:\OneDrive\br-data
python app.py
```

#### 10.4 Data source endpoints

| Variable | Description |
| --- | --- |
| `BR_API_BASE` | BR detail endpoint; `{}` is replaced with the BR ID |
| `KM_API_BASE` | Killmail endpoint; `{}` is replaced with the killmail ID |
| `ESI_BASE` | Official EVE ESI root, used by `ship_id.py` |
| `BR_LINK_PATTERN` | BR link validation regex, **shared by client and server** |
| `REQUEST_TIMEOUT` | Per-request timeout in seconds; raise to 20–30 on slow links |

`BR_LINK_PATTERN` is special: it is used by `re.match` in Python **and** injected into the
front-end JavaScript via `| tojson` as `new RegExp(...)`.
**Change it once and both sides stay in sync** — no drift between validators.

#### 10.5 Report output

| Variable | Default | Description |
| --- | --- | --- |
| `REPORT_FILE_PREFIX` | `战斗报告` | Report file name prefix |
| `CHART_TOP_N` | `10` | How many entries the charts show |

Resulting file name looks like: `战斗报告_20250925_test.xlsx`

#### 10.6 Attendance

| Variable | Default | Description |
| --- | --- | --- |
| ★ `EVENT_TYPES` | `["反收割", "小队活动", "会战", "集结待命"]` | Options in the "event type" dropdown |

Customise to match your own corporation's terminology:

```python
EVENT_TYPES = ["Counter-roam", "Small gang", "Fleet fight", "Staging"]
```

### 11. What you must edit yourself

**How to locate them** — run this from the project root:

```bash
grep -rn "TODO(通用版)\|【需修改】" --include="*.py" --include="*.html" --include="*.txt" .
```

> Line numbers below reflect the initial v1.1 release and may drift after edits —
> **always trust the search results over the numbers.**

#### 11.1 Must change (affects correctness)

| # | Location | File | Line | What to do |
| --- | --- | --- | --- | --- |
| ① | Web UI → Corporation | — | — | Add your corporation ID and name, or nothing will be counted |
| ② | Branding | `config.py` | 24 | `SITE_NAME` → your system name |
| ③ | Branding | `config.py` | 30 | `FOOTER_TEXT` → your footer |
| ④ | Web service | `config.py` | 55 | `SECRET_KEY` → a random string (mandatory if public) |

#### 11.2 Recommended

| # | File | Line | What to do |
| --- | --- | --- | --- |
| ⑤ | `config.py` | 34 | `CORP_LOGO` → your logo, or `""` to hide it |
| ⑥ | `config.py` | 38 | `BACKGROUND_IMAGE` → your background |
| ⑦ | `config.py` | 46 / 49 | `HOST` / `PORT` |
| ⑧ | `config.py` | 125 | `EVENT_TYPES` — attendance event types |
| ⑨ | `version_info.txt` | from line 4 | Company name / copyright shown in EXE properties |
| ⑩ | `build.py` | 27 / 31 / 34 / 37 | EXE name, icon, app name, version-info file |

#### 11.3 Optional

| File | Line / location | Description |
| --- | --- | --- |
| `config.py` | 64 | `DATA_FOLDER` — custom data directory |
| `config.py` | 94 / 97 / 100 | Endpoint URLs |
| `config.py` | 103 | `BR_LINK_PATTERN` validation rule |
| `config.py` | 106 | `REQUEST_TIMEOUT` |
| `config.py` | 114 | `REPORT_FILE_PREFIX` |
| `config.py` | 117 | `CHART_TOP_N` |
| `attendance.py` | line 25, `STORE_FILENAME` | Attendance data file name |
| `templates/base.html` | around line 13 | CDN URLs (swap for local files on an isolated network) |
| `templates/attendance.html` | 24 | Event type dropdown |
| `tools/make_icon.py` | 22 | Icon colours |
| `app.py` | 316 | "No corporation configured" message text |
| `app.py` | 474 | Character management extension point (whitelists, etc.) |

#### 11.4 Asset replacement cheat sheet

| Want to change | Steps |
| --- | --- |
| Navbar logo | Put image in `static/` → set `config.CORP_LOGO` to the file name |
| Page background | Put image in `static/` → set `config.BACKGROUND_IMAGE` |
| Favicon + EXE icon | Edit colours in `tools/make_icon.py` → `python tools/make_icon.py`; or replace `static/favicon.ico` and `app_icon.ico` directly |
| Attendance event types | Edit `config.EVENT_TYPES` |

---

## Part 3 · Feature Walkthroughs

### 12. Combat statistics

Entry point: the home page (navbar → "战斗统计分析").

#### 12.1 Form fields

| Field | Required | Description |
| --- | --- | --- |
| 选择报告日期 — report date | ✅ | Used in the file name, format `YYYYMMDD` |
| 自定义报告名称 — custom name | ✅ | File name suffix, max 50 chars; illegal characters become `_` |
| 粘贴BR链接 — BR links | ✅ | One per line; must be `https://br.evetools.org/br/<ID>` |

#### 12.2 Steps

1. Pick a date and type a name;
2. Paste BR links (one per line) — below the box a live counter shows
   "共 N 条链接，格式全部正确 / 其中 M 条无效" (N links, all valid / M invalid);
3. Click **生成报告 (Generate report)** — the progress bar advances through
   *initialising → analysing combat data → generating the report*, then the page reloads;
4. When the green banner *"BR报告生成完毕"* appears, click **下载最新报告 (Download latest report)**.

#### 12.3 Counting rules

Only members of corporations listed under Corporation management are counted.
For each killmail:

- A friendly pilot appears in the **attacker** list → their ship and damage are recorded;
- A friendly pilot is the **victim** → a sortie and their ship are recorded;
- An attacker with `blow = true` (final blow) → that pilot's kill count +1;
- Within one fight, a pilot's sortie is counted once.

#### 12.4 Output

See [Chapter 15](#15-report-files-explained).

### 13. Attendance statistics

#### 13.1 Important note about the URL

Because of Flask blueprint prefix stacking (`url_prefix='/attendance'` plus
`@route('/attendance')` inside the blueprint), the attendance page's **real URL is**:

```
http://127.0.0.1:5000/attendance/attendance
```

**Opening `/attendance` directly returns 404. This is expected behaviour, not a bug.**
Normally you just click "出勤统计分析" in the navbar.

#### 13.2 Form fields

| Field | Required | Description |
| --- | --- | --- |
| 报告日期 — report date | ✅ | Defaults to today |
| 报告类型 — event type | ✅ | Options from `config.EVENT_TYPES` |
| 舰队说明 — note | ❌ | Free text, e.g. "Hauling escort" |
| 角色列表 — member list | ✅ | One character name per line |

#### 13.3 Steps

1. Pick a date and an event type;
2. Paste the members who joined (one per line);
3. Click **添加舰队角色 (Add fleet members)** → a modal reports "已成功添加 N 个角色",
   the textarea clears, and focus returns to it;
4. The right-hand table **refreshes live** with the attendance ranking; the footer shows the total;
5. Repeat 1–4 for the next fleet;
6. Click **导出Excel (Export Excel)** when you want an archive.

**Deduplication**: a name repeated inside a single submission counts once.
So "Pilot Alpha / Pilot Beta / Pilot Alpha" will report 2 members added.

#### 13.4 Persistence

Attendance lives in `data/attendance.json` — **close and reopen the app, data is still there**.
On page load the app checks for existing data and, if found, enables the export button and
refreshes the table immediately.

### 14. Admin pages

Navbar → **管理 (Manage)** dropdown.

#### 14.1 Corporation management

- **Purpose**: maintain the list of corporations to track
- **File**: `data/corp_id.txt`
- **Note**: empty by default in the Generic Edition; you must add at least one

Fields: corporation ID (digits) and corporation name (display only).
Existing entries are listed with a count and can be deleted individually (with confirmation).

#### 14.2 Character management

- **Purpose**: supplement or correct the character ID ↔ name mapping
- **File**: `data/char_id.txt`

> **Usually no manual work needed.** During combat statistics, the app automatically appends
> every friendly character it encounters in the BRs into `char_id.txt`.
>
> You only need this page when a report shows `Unknown(1234567890)`, or when you want to
> pre-seed a name for someone who has not appeared yet.

#### 14.3 Report history

- **Purpose**: view, download, and delete generated reports
- **Location**: `data/reports/`
- Sorted newest first; deletion asks for confirmation

#### 14.4 System log

- **Purpose**: troubleshooting
- **File**: `data/system.log`
- Records every `SUCCESS` / `ERROR` / `INFO` entry with a timestamp

Example:

```
[2025-09-25 10:12:31.220000] SUCCESS: Report generated: 战斗报告_20250925_test.xlsx
[2025-09-25 10:15:02.100000] ERROR: Invalid BR link: https://br.evetools.org/br/xxx
```

A more verbose runtime log is written to `data/detailed.log` by Python's `logging` module.

### 15. Report files explained

#### 15.1 Combat statistics report

File name: `战斗报告_<date>_<custom name>.xlsx`
Default location: `data/reports/`

**Sheet "战斗统计"**

| Col | Field | Description |
| --- | --- | --- |
| A | Rank | Descending by total damage |
| B | Pilot | Character name |
| C | Sorties | Times the pilot appeared in the combat records |
| D | Kills | Final blows |
| E | Total damage | Accumulated damage |
| F | Ships | Deduplicated ship list, separated by an ideographic comma |

**Charts**

- At `H2`: "总伤害排名前10" — Total damage top 10 (column chart)
- At `H18`: "击杀数排名前10" — Kills top 10 (column chart)

How many entries the charts show is controlled by `config.CHART_TOP_N`.

#### 15.2 Attendance report

File name: `舰队出勤统计_<timestamp>.xlsx`
Downloaded to your browser's default download folder.

| Sheet | Contents |
| --- | --- |
| 出勤统计 | Character name, attendance count (descending) |
| 历史记录 | Submission time, event date, event type, member count, note, member list |

---

## Part 4 · Packaging and Publishing

### 16. Building an EXE

Useful when handing the tool to teammates who do not have Python.

```bash
# 1) Install build dependencies
pip install pyinstaller pillow

# 2) (Optional) regenerate the icon
python tools/make_icon.py

# 3) Build
python build.py
```

Output: `dist/BRStatisticsTool.exe`

Build configuration sits at the top of `build.py`:

| Constant | Description |
| --- | --- |
| `EXE_NAME` | Output file name |
| `ICON_FILE` | Icon file name (`.ico`) |
| `APP_NAME` | Application name shown in Windows properties (defaults to `config.SITE_NAME`) |
| `VERSION_FILE` | Version resource file (default `version_info.txt`) |

The build script automatically:

1. Cleans `build/`, `dist/`, and any stale `.spec` file;
2. Bundles `templates/`, `static/`, and `data/` into the EXE;
3. Declares hidden imports (`jinja2`, `werkzeug`, `flask`, `openpyxl`, `pandas`, …);
4. Applies the icon and version information.

**First launch of the EXE**: Windows SmartScreen may warn you — click
"More info → Run anyway".

> 💡 After packaging, **data is written to the `data/` folder next to the EXE**, not to a
> temp directory — so nothing is lost when you close the app. This was one of the bugs
> fixed in the Generic Edition.

### 17. Distribution and migration

When sharing, copy:

```
BRStatisticsTool.exe
data/                    # ship_id.txt alone is enough
```

**Migrating to a new machine**: copy the whole `data/` directory to the same relative
location on the new machine. Historical reports and attendance records travel with it.

### 18. Publishing to GitHub

The project ships with a `.gitignore` that already excludes:

- `build/`, `dist/`, `*.spec` (build artifacts)
- `data/corp_id.txt`, `data/char_id.txt`, `data/br_links.txt` (your data)
- `data/attendance.json`, `data/*.log`, `data/reports/*.xlsx`
- `.venv/`, `.env`, `*.pem`, `*.key`

Standard flow:

```bash
git init
git add .
git commit -m "feat: generic edition"
git branch -M main
git remote add origin https://github.com/<user>/<repo>.git
git push -u origin main
```

#### ⚠️ Token security

**Never put a token in code, scripts, `.git/config`, or a commit.**

If a token ever appeared in a chat log, terminal output, or a screenshot, immediately:

1. Open GitHub → `Settings → Developer settings → Personal access tokens`
2. Find the token → **Revoke**
3. Generate a new one with only the scopes you need (`Contents: Read and write`)

After pushing, verify nothing lingered in `.git/config`:

```bash
grep -c "github_pat\|ghp_" .git/config || echo "clean"
```

#### If `git push` fails

Errors such as `schannel: server closed abruptly`, `SSL_ERROR_SYSCALL`, or a long hang
usually mean **the git endpoint on `github.com` is blocked by your network**
(while `api.github.com` still works). In that case push via the GitHub REST API —
see the bundled skill `github-push-via-api`.

---

## Part 5 · Operations and Troubleshooting

### 19. Backup and restore

#### What to back up

| File | Priority | Notes |
| --- | --- | --- |
| `data/attendance.json` | 🔴 High | Attendance records — painful to re-enter |
| `data/corp_id.txt` | 🔴 High | Corporation configuration |
| `data/char_id.txt` | 🟡 Medium | Character mapping; can be rebuilt |
| `data/reports/*.xlsx` | 🟡 Medium | Historical reports |
| `data/ship_id.txt` | 🟢 Low | Public data; re-fetchable |

**Simplest approach**: copy the entire `data/` directory.

**Automatic backup**: point the data directory at a cloud-synced folder:

```bash
set BR_DATA_DIR=D:\OneDrive\br-data
python app.py
```

#### Restore

Copy the backed-up `data/` back into place. Nothing else to do.

### 20. Upgrade procedure

1. Back up `data/` (previous chapter);
2. Overwrite everything **except `data/`** with the new version;
3. **If `config.py` changed upstream**, re-apply your own values;
4. Re-run `pip install -r requirements.txt` (dependencies may have changed);
5. Start it and run `python tools/smoke_test.py` to confirm everything works.

### 21. Error message reference

| Page message / console output | Cause | Fix |
| --- | --- | --- |
| `Invalid links: xxx` | BR link format invalid | Must be `https://br.evetools.org/br/<ID>` with no extra parameters |
| `Invalid BR link: xxx` | Format fine but endpoint returned non-200 | BR expired or endpoint temporarily down — retry |
| `Failed to access: xxx` | Same as above | Same as above |
| `尚未配置任何军团 ID，请先到「军团管理」页面添加…` | `corp_id.txt` is empty | Add a corporation under Corporation management |
| `Date and name are required` | Date or name missing | Both are mandatory |
| `Please enter valid BR links` | Links box empty | Enter at least one link |
| `Report generation failed: xxx` | Unexpected error during generation | Check the system log for details |
| `[提示] 数据文件 corp_id.txt 不存在，已按空数据处理` | First run; file not created yet | **Normal.** Created automatically once you configure a corporation |
| `Unknown(2118200607)` in a report | Character mapping missing | Auto-completed next run; or add it under Character management |
| `Unknown(12345)` for a ship | New ship not in the mapping | Run `python ship_id.py` |
| Requests keep timing out | Network / proxy issue | Turn off VPN; or raise `REQUEST_TIMEOUT` |
| 404 at `/attendance` | Blueprint prefix stacking | Use `/attendance/attendance` or click through the navbar |
| Port already in use | Another process holds 5000 | Change `config.PORT` |

### 22. FAQ

**Q1: Clicking attendance shows `/attendance` and a 404 page?**

Blueprint prefix stacking — the real URL is `/attendance/attendance`.
Clicking through the navbar always works. See [13.1](#131-important-note-about-the-url).

**Q2: Why does the very first report fail saying no corporation is configured?**

The Generic Edition ships with an empty corporation list. This is deliberate — otherwise it
would silently produce an all-blank report and you would assume the tool was broken.
Add one under Manage → Corporation.

**Q3: Can I track several corporations at once?**

Yes. Add each one under Corporation management. The scope is the union of all configured
corporations.

**Q4: Do I have to paste BR links one by one?**

No — paste them as multiple lines, one per line. The app processes them in sequence and
tells you exactly which link is malformed.

**Q5: Why do the damage numbers differ from in-game figures?**

This tool reports **damage recorded in the BRs you supplied**, not the in-game damage
statistics. Differences come from: only the links you provided are covered; kills only count
final blows; and within one fight a pilot's sortie counts once.

**Q6: Can attendance data be lost?**

No. It is stored in `data/attendance.json` and survives restarts. Back that file up.

**Q7: How do others on my LAN access it?**

Set `HOST = "0.0.0.0"` and have them browse to `http://<your-LAN-IP>:5000`.
Allow the port through the firewall. **Change `SECRET_KEY` before exposing it.**

**Q8: Can I host it on the public internet?**

Technically yes, but **the app has no authentication** — public exposure means anyone can
use it. If you must, add authentication first
(see [24.4](#244-how-do-i-add-login-authentication)) and put it behind HTTPS.

**Q9: A ship shows up as a number in the report — why?**

New ship not yet in the mapping. Run `python ship_id.py`; the script writes directly to
`data/ship_id.txt` by default.

**Q10: Can I replace the background and logo?**

Yes. Put the image in `static/`, then set `config.CORP_LOGO` / `config.BACKGROUND_IMAGE`.
Make sure you have the rights to the artwork you use.

**Q11: Does it run on macOS / Linux?**

Yes, it runs. Note that `build.py` produces a Windows EXE; on Linux/macOS it produces a
native binary for that platform instead.

**Q12: Can it run in the background without a console window?**

On Windows you can use `pythonw app.py` (no console window), or add the packaged EXE to
your Startup folder. Bear in mind you lose the live log output — for troubleshooting,
running it from a terminal is still more convenient.

---

## Part 6 · Developer Guide

### 23. Code structure and key functions

#### 23.1 app.py

| Function / route | Purpose |
| --- | --- |
| `inject_globals()` | Injects branding and navigation into every template |
| `data_path(*parts)` | Builds paths inside the data directory (**always use this, never `__file__`**) |
| `load_ids(filename)` | Reads `ID_name` mappings; skips blank lines and `#` comments |
| `save_ids(filename, data)` | Writes mappings back |
| `resolve_report_path(filename)` | **Security check**: only allows real `.xlsx` files inside `data/reports` |
| `process_killmail(...)` | Parses one killmail and accumulates the numbers |
| `generate_final_data(...)` | Aggregates into the final report structure |
| `/` | Home page: submit BR links, generate a report |
| `/download-report/<filename>` | Download (with whitelist validation) |
| `/delete-report/<filename>` | Delete (with whitelist validation) |
| `/manage-chars`, `/manage-corps` | Character / corporation management |
| `/history`, `/logs`, `/docs` | History / log / help pages |

#### 23.2 attendance.py

| Function | Purpose |
| --- | --- |
| `_store_path()` | Absolute path of the attendance data file |
| `_load_submissions()` | Reads all records (returns `[]` if the file is missing or corrupt) |
| `_save_submissions()` | Writes all records back |
| `_build_stats()` | Aggregates and sorts attendance counts by name |
| `/submit_fleet` | Submit a member list (**deduplicated on submit**) |
| `/get_stats`, `/get_history` | Live stats / submission history |
| `/export_stats` | Excel export via an in-memory `BytesIO` stream — nothing hits disk |

#### 23.3 Three security-relevant implementations (**do not regress these**)

1. **Path traversal protection** — `resolve_report_path()`
   Accepts only file names that contain no path separator, end in `.xlsx`, and appear in the
   actual listing of `data/reports`.
   > ⚠️ **Do not switch to `werkzeug.secure_filename`** — it strips non-ASCII characters
   > entirely and would break Chinese report file names, making downloads fail.

2. **XSS protection** — `escapeHtml()` in `templates/attendance.html`
   Character names come from user input; always escape before writing to the DOM.

3. **Single source of truth for validation** — `config.BR_LINK_PATTERN` is injected into
   JavaScript via `| tojson`, so client and server can never disagree.

### 24. Extending the system

#### 24.1 Adding a new configuration option

1. Add the variable to the right section of `config.py`, with a `【需修改】` marker or a
   clear comment;
2. If templates need it, expose it in `inject_globals()` in `app.py`;
3. Update Chapter 10 / 11 of `操作手册.md` **and** the matching chapter in this file.

#### 24.2 Adding a new page

```python
@app.route('/my-page')
def my_page():
    return render_template('my_page.html', ...)
```

Inherit from `base.html` to get the navbar, footer, and styling for free:

```jinja
{% extends "base.html" %}
{% block content %}
  ...your content...
{% endblock %}
```

Register the nav entry in `inject_globals()` in `app.py`.

#### 24.3 Changing what gets counted

For example, to count "any participation in a kill" instead of only final blows, edit the
final-blow block in `process_killmail()`:

```python
# Current behaviour: only blow=true increments the counter
for attacker in kb_data.get('atts', []):
    if attacker.get('blow', True):
        ...
        char_dir[char_str][0][0] += 1
        break
```

Drop the `blow` check and the `break` to count every participation — but judge for yourself
whether that metric is meaningful for your use case.

#### 24.4 How do I add login authentication?

The original code left a comment about a login endpoint but never implemented it.
A reasonable approach:

1. `pip install flask-login`;
2. Add a `LoginManager` in `app.py`;
3. Keep a simple user store (a JSON file under `data/` is fine to start);
4. Decorate your routes with `@login_required`.

**Do not** hardcode usernames and passwords in the source — that is the first thing scanners
look for.

#### 24.5 Always verify after changing code

```bash
python -m py_compile app.py config.py attendance.py    # syntax check
python tools/smoke_test.py                             # 25 smoke assertions
```

`smoke_test.py` points the data directory at a temporary folder, so it
**never touches your real `data/`** — run it as often as you like.

#### 24.6 Translating the UI into another language

The templates are plain Jinja2 HTML. To localise:

1. Replace the Chinese strings in `templates/*.html`
   (the nav labels, buttons, and help text are the main ones);
2. Update `config.SITE_NAME`, `SITE_VERSION`, `FOOTER_TEXT`, and `EVENT_TYPES`;
3. Update `templates/docs.html` (the built-in help page).

Avoid touching the `url_for(...)` endpoint names — those are code identifiers, not display
strings, and renaming them breaks routing.

### 25. Generic edition changelog

#### 25.1 Content removed from the original customised build

| Item | How it was handled |
| --- | --- |
| Corporation logo `static/corp_logo.png` | Deleted; replaced with the neutral `site_logo.svg` |
| Corporation icon `cosmic_wanderers_icon.ico` | Deleted; replaced with `app_icon.ico` / `favicon.ico` |
| Corporation background (2.3 MB of official art) | Deleted; replaced with a CSS deep-space gradient |
| Corporation name `Cosmic-Wanderers` | Replaced everywhere with the configurable `config.SITE_NAME` |
| Corporation ID | `data/corp_id.txt` cleared |
| 260 real character IDs / names | `data/char_id.txt` cleared |
| 3 real BR links | `data/br_links.txt` cleared |
| Combat and report logs | `system.log`, `detailed.log` deleted |
| Author credits | Removed |
| Leftovers `output.xlsx`, `__pycache__/` | Deleted |

#### 25.2 Bugs fixed

| # | Problem | Fix |
| --- | --- | --- |
| 1 | Attendance lived only in memory and was lost on restart | Persisted to `data/attendance.json` |
| 2 | Excel export left an `output.xlsx` in the working directory | Switched to an in-memory `BytesIO` stream |
| 3 | Report endpoints were vulnerable to path traversal | Added a whitelist check |
| 4 | Attendance table injected character names into `innerHTML` (XSS) | Added `escapeHtml()` |
| 5 | Packaged EXE wrote data into a temp directory, losing it on restart | Data directory now resolves next to the EXE; `BR_DATA_DIR` supported |
| 6 | File I/O did not specify an encoding | Standardised on `encoding='utf-8'` |
| 7 | Client and server each had their own BR validation regex | Unified via `config.BR_LINK_PATTERN` |
| 8 | Four admin pages had no navigation entry | Added a "Manage" dropdown to the navbar |
| 9 | Form contents were lost after a validation error | Date, name, and links are now re-displayed; last links saved |
| 10 | Bootstrap loaded twice (5.1.3 + 5.3.0), Font Awesome duplicated | Consolidated to a single version on bootcdn |
| 11 | Empty report generated silently when no corporation was configured | Now fails with a clear message pointing to the config page |
| 12 | Duplicate names in one attendance list were double counted | Deduplicated on submit, order preserved |
| 13 | Delete button built `onclick` by string concatenation (broke on quotes) | Now uses `data-*` attributes plus `encodeURIComponent` |
| 14 | Empty lists rendered as a blank page | Added empty states and item counts |
| 15 | Hardcoded personal absolute paths | Replaced with `config.get_data_folder()` |
| 16 | Malformed markup in `docs.html` (nested `<h3>`, unclosed `<ol>`) | Rewritten as valid HTML |

#### 25.3 Added

- `config.py` — single configuration entry point
- `README.md` — project overview
- `操作手册.md` — detailed manual (Chinese)
- `OPERATION_MANUAL_EN.md` — detailed manual (English, this file)
- `data/README.md` — data file format reference
- `tools/make_icon.py` — icon generator
- `tools/smoke_test.py` — smoke test (25 assertions)
- `.gitignore`, `requirements.txt`
- Neutral logo / favicon / EXE icon

---

## Appendices

### Appendix A: Data file formats

Convention across all files: **UTF-8, one record per line, ID and name separated by the
first underscore `_`.** Blank lines and lines starting with `#` are ignored.

#### A.1 `corp_id.txt`

```
# corporationID_corporationName
98769610_Example Corporation
12345678_Another Corporation
```

#### A.2 `char_id.txt`

```
2118200607_Character Name
2113436255_Another Pilot
```

#### A.3 `ship_id.txt`

```
582_Bantam
583_Condor
584_Griffin
```

Public data; 6000+ entries preloaded. Refresh with `python ship_id.py` when new ships appear.

#### A.4 `br_links.txt`

```
https://br.evetools.org/br/68d956f22213cf0012ace2af
https://br.evetools.org/br/68d6bab14ff6700012d935c5
```

Used to restore the last submitted links into the textarea.

#### A.5 `attendance.json`

```json
[
  {
    "timestamp": "2025-09-25T10:12:31.220000",
    "date": "2025-09-25",
    "event_type": "Small gang",
    "note": "Hauling escort",
    "members": ["Pilot Alpha", "Pilot Beta"],
    "member_count": 2
  }
]
```

> Do not corrupt this file manually — a malformed file is treated as "no attendance data".

### Appendix B: Configuration cheat sheet

```python
# ---- Branding ----
SITE_NAME         = "EVE Online BR 战斗统计系统"
SITE_VERSION      = "1.1"
FOOTER_TEXT       = "EVE Online BR 战斗统计系统"
CORP_LOGO         = "site_logo.svg"        # "" = hide
BACKGROUND_IMAGE  = ""                     # "" = built-in gradient

# ---- Service ----
HOST              = "127.0.0.1"            # "0.0.0.0" = LAN-reachable
PORT              = 5000
AUTO_OPEN_BROWSER = True
SECRET_KEY        = "please-change-this-secret-key"   # must change if public

# ---- Data directory ----
DATA_FOLDER       = ""                     # "" = default; BR_DATA_DIR env var wins

# ---- Endpoints ----
BR_API_BASE       = "https://br.evetools.org/api/v1/composition/get/{}"
KM_API_BASE       = "https://kb.evetools.org/api/v1/killmails/{}"
ESI_BASE          = "https://esi.evetech.net/latest"
BR_LINK_PATTERN   = r"^https://br\.evetools\.org/br/[\w-]+$"
REQUEST_TIMEOUT   = 10

# ---- Reports ----
REPORT_FILE_PREFIX = "战斗报告"
CHART_TOP_N        = 10

# ---- Attendance ----
EVENT_TYPES = ["反收割", "小队活动", "会战", "集结待命"]
```

### Appendix C: Glossary

| Term | Full form | Meaning |
| --- | --- | --- |
| BR | Battle Report | A link recording the participants and outcome of a fight |
| KM | Killmail | The record generated for every kill in EVE |
| Corporation | — | The basic organisational unit in EVE |
| Alliance | — | A larger organisation composed of multiple corporations |
| Sortie | — | How many times a pilot appears in the combat records |
| Final blow | — | The last hit of a kill; marked `blow = true` in a BR |
| Ship | — | A flyable vehicle in EVE |
| ESI | EVE Swagger Interface | EVE's official public API |
| Blueprint | — | The Flask component that groups routes under a URL prefix |

---

### Quick start, in one minute

```
1. pip install -r requirements.txt
2. python app.py
3. Browser → Manage → Corporation → add "<corpID>_<corpName>"
4. Home → fill in date and name → paste BR links → Generate report → Download
5. Attendance → navbar "出勤统计分析" → enter members → Export Excel
```

---

> This document mirrors [操作手册.md](#中文操作手册) chapter for chapter.
> If the two ever disagree, **the Chinese version is authoritative** — it is the
> primary maintained document.

---

<sub>
本 README 由 <code>tools/build_readme.py</code> 从 <a href="操作手册.md">操作手册.md</a> 与
<a href="OPERATION_MANUAL_EN.md">OPERATION_MANUAL_EN.md</a> 自动生成，请勿手工编辑——
修改文档请改上述两份源文件后重新执行该脚本。<br>
This README is generated by <code>tools/build_readme.py</code> from the two manual source
files above. Do not edit it by hand.
</sub>
