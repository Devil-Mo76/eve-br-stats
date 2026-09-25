# EVE Online BR 战斗统计系统（通用版）

一个用于分析 **EVE Online** 战斗报告（BR）的本地统计工具：粘贴 BR 链接，自动汇总参战成员的
**出击次数 / 击杀数 / 总伤害 / 使用舰船**，导出带柱状图的 Excel 报表；同时提供舰队**出勤统计**功能。

> 本仓库是**通用版本**：不含任何特定军团、联盟或个人的数据与标识，首次使用只需在网页上添加
> 你自己的军团 ID 即可。

---

## 功能一览

| 模块 | 说明 |
| --- | --- |
| 战斗统计分析 | 解析多个 BR 链接，按「军团 ID」筛选成员，统计出击次数、击杀数、总伤害、使用舰船 |
| 出勤统计分析 | 按活动逐次录入舰队成员名单，实时统计出勤次数，可导出 Excel |
| 报表导出 | Excel 含数据表 + 「总伤害 Top10」「击杀数 Top10」两张柱状图 |
| 管理后台 | 军团管理 / 角色管理 / 历史报告 / 系统日志 |

---

## 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动（默认 http://127.0.0.1:5000，会自动打开浏览器）
python app.py
```

浏览器中打开后，先进入 **管理 → 军团管理** 添加你要统计的军团，然后回到首页提交 BR 链接即可。

---

## 目录结构

```
br-web/
├── app.py                  # 主程序（路由 + 战斗统计核心逻辑 + Excel 生成）
├── config.py               # ★ 通用版配置入口（品牌 / 端口 / 接口 / 出勤选项）
├── attendance.py           # 出勤统计蓝图（数据落盘 data/attendance.json）
├── ship_id.py              # 舰船 ID→名称 映射更新脚本（EVE ESI）
├── build.py                # PyInstaller 打包脚本
├── version_info.txt        # EXE 版本信息（公司名 / 说明 / 版权）
├── requirements.txt
├── 操作手册.md             # ★ 完整操作手册 + 需自行修改清单
├── app_icon.ico            # EXE 图标（由 tools/make_icon.py 生成）
├── tools/
│   └── make_icon.py        # 图标生成脚本（favicon + exe 图标 + PNG）
├── data/                   # 运行时数据目录（详见 data/README.md）
│   ├── ship_id.txt         # 舰船映射（公共数据，已预置）
│   ├── reports/            # 生成的 Excel 报表
│   └── README.md
├── static/
│   ├── site_logo.svg       # 导航栏 LOGO（中性占位，可替换）
│   ├── favicon.ico         # 浏览器图标
│   └── style.css
└── templates/              # Jinja2 模板
```

---

## 配置与二次开发

**所有需要你修改的地方都集中在两个位置：**

1. **`config.py`** —— 品牌名、LOGO、背景图、端口、SECRET_KEY、数据源接口、出勤活动类型等，
   全部标注为 `【需修改】`。
2. **源码中的 `TODO(通用版)` 注释** —— 在编辑器里全局搜索 `TODO(通用版)` 或 `【需修改】`，
   即可逐条定位。

完整清单见 **[操作手册.md](操作手册.md)** 的「需自行修改清单」章节。

---

## 打包为 EXE

```bash
pip install pyinstaller
python build.py
# 产物：dist/BRStatisticsTool.exe
```

打包后数据写在 **EXE 同级目录** 的 `data/` 下，连同 `data/` 一起拷贝即可迁移。

---

## 数据来源

- BR 数据：<https://br.evetools.org>
- 击杀邮件：<https://kb.evetools.org>
- 舰船数据：<https://esi.evetech.net>

本工具只做数据聚合，不对数据源内容的准确性负责。

## 免责声明

EVE Online 及相关素材版权归 CCP hf. 所有。本项目为非官方的第三方工具，
未使用任何官方美术素材，与 CCP 无隶属关系。
