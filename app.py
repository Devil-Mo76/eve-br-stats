# -*- coding: utf-8 -*-
"""
EVE Online BR 战斗统计系统 · 通用版  主程序
--------------------------------------------------------------------------
启动方式：
    源码运行   python app.py
    打包运行   python build.py  （生成 dist/BRStatisticsTool.exe）
--------------------------------------------------------------------------
通用版改动说明：
    1. 全部品牌信息（系统名 / LOGO / 页脚 / 背景）改为从 config.py 读取，代码中不再硬编码；
    2. 数据目录支持环境变量 BR_DATA_DIR 覆盖，打包成 exe 后数据写在 exe 同级目录；
    3. 下载 / 删除报告接口增加路径白名单校验，防止目录穿越；
    4. 文件读写统一指定 UTF-8 编码；
    5. 补齐导航入口（军团管理 / 角色管理 / 历史报告 / 系统日志）。
--------------------------------------------------------------------------
寻找需要自行修改的位置：全项目搜索  TODO(通用版)  或  【需修改】
"""

import logging
import os
import re
from datetime import datetime

import requests
from flask import (Flask, jsonify, redirect, render_template, request,
                   send_file, url_for)
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, Side

import config as cfg
from attendance import attendance_bp

app = Flask(__name__)
app.config['SECRET_KEY'] = cfg.SECRET_KEY
app.config['DATA_FOLDER'] = cfg.get_data_folder()
app.register_blueprint(attendance_bp, url_prefix='/attendance')


# ============================== 模板全局变量 ==============================

@app.context_processor
def inject_globals():
    """把所有模板都要用的品牌信息、导航项注入到 Jinja 上下文"""
    return {
        'site_name': cfg.SITE_NAME,
        'site_version': cfg.SITE_VERSION,
        'footer_text': cfg.FOOTER_TEXT,
        'corp_logo': cfg.CORP_LOGO,
        'background_image': cfg.BACKGROUND_IMAGE,
        'br_link_pattern': cfg.BR_LINK_PATTERN,
        'current_year': datetime.now().year,
        'nav_items': [
            {'name': '战斗统计分析', 'url': url_for('index'), 'icon': 'fa-home'},
            {'name': '出勤统计分析', 'url': url_for('attendance.attendance_index'), 'icon': 'fa-users'},
            {'name': '操作文档', 'url': url_for('docs'), 'icon': 'fa-file-alt'},
        ],
        'manage_items': [
            {'name': '军团管理', 'url': url_for('manage_corps'), 'icon': 'fa-shield-halved'},
            {'name': '角色管理', 'url': url_for('manage_chars'), 'icon': 'fa-user-gear'},
            {'name': '历史报告', 'url': url_for('history'), 'icon': 'fa-clock-rotate-left'},
            {'name': '系统日志', 'url': url_for('logs'), 'icon': 'fa-list-ul'},
        ],
    }


# ================================ 工具函数 ================================

def data_path(*parts):
    """拼接数据目录下的文件路径"""
    return os.path.join(app.config['DATA_FOLDER'], *parts)


def load_ids(filename):
    """
    加载「ID_名称」格式的映射文件，返回 {id: name}。
    空行与以 # 开头的注释行会被忽略，方便手工编辑数据文件。
    """
    id_dict = {}
    try:
        with open(data_path(filename), 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                if '_' in line:
                    idnum, name = line.split('_', 1)
                    id_dict[idnum.strip()] = name.strip()
    except FileNotFoundError:
        # 数据文件缺失不是错误：通用版首次运行就是空的，允许用户从「管理」页面添加
        print(f"[提示] 数据文件 {filename} 不存在，已按空数据处理")
    return id_dict


def save_ids(filename, data):
    """保存「ID_名称」格式的映射文件"""
    path = data_path(filename)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        for idnum, name in data.items():
            f.write(f"{idnum}_{name}\n")


def load_br_links():
    """读取上次保存的 BR 链接"""
    try:
        with open(data_path('br_links.txt'), 'r', encoding='utf-8') as f:
            return [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        return []


def extract_br_id(url):
    """从 BR 链接中提取 BR ID"""
    return url.rstrip('/').split('/')[-1]


def log_event(status, message):
    """写入系统日志"""
    log_entry = f"[{datetime.now()}] {status}: {message}\n"
    path = data_path('system.log')
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'a', encoding='utf-8') as f:
        f.write(log_entry)


def resolve_report_path(filename):
    """
    校验报告文件名并返回其绝对路径。
    只允许访问 data/reports 下真实存在、且以 .xlsx 结尾的文件，
    防止通过 ../../ 之类的路径穿越读写到系统其它文件。
    """
    reports_dir = data_path('reports')
    # 文件名里不允许出现任何路径分隔符
    if not filename or filename != os.path.basename(filename):
        return None
    if not filename.endswith('.xlsx'):
        return None
    if not os.path.isdir(reports_dir):
        return None
    if filename not in os.listdir(reports_dir):
        return None
    return os.path.join(reports_dir, filename)


# ================================ 核心处理 ================================

def process_killmail(km, corp_ids, char_id, char_dir, index):
    """处理单条击杀邮件，把本军团成员的出击 / 击杀 / 伤害累加进 char_dir"""
    kill_id = km['id']
    response_kb = requests.get(cfg.KM_API_BASE.format(kill_id), timeout=cfg.REQUEST_TIMEOUT)

    if response_kb.status_code != 200:
        return

    kb_data = response_kb.json()
    all_name = kb_data.get('names', {}).get('chars', {})

    # 补全攻击者名称
    for d in kb_data.get('atts', []):
        corp_str = str(d.get('corp', ''))
        char_str = str(d.get('char', ''))
        if corp_str in corp_ids and char_str not in char_id:
            char_id[char_str] = all_name.get(char_str, 'Unknown Character')
            save_ids('char_id.txt', char_id)

    # 补全受害者名称
    if 'victim' in kb_data:
        vict = kb_data['victim']
        vict_corp = str(vict.get('corp', ''))
        vict_char = str(vict.get('char', ''))
        if vict_corp in corp_ids and vict_char not in char_id:
            char_id[vict_char] = all_name.get(vict_char, 'Unknown Character')
            save_ids('char_id.txt', char_id)

    # 统计受害者数据（被击毁 = 一次出击）
    if 'victim' in kb_data and 'corp' in kb_data['victim']:
        if str(kb_data['victim']['corp']) in corp_ids:
            char_key = str(kb_data['victim']['char'])
            entry = char_dir.setdefault(char_key, [[0], [1], True])
            entry.extend([[kb_data['victim']['ship'], 0]])

            if index != 0 and not entry[2]:
                entry[1][0] += 1
                entry[2] = True

    # 统计攻击者数据
    for attacker in kb_data.get('atts', []):
        if str(attacker.get('corp', '')) in corp_ids:
            char_key = str(attacker.get('char', ''))
            entry = char_dir.setdefault(char_key, [[0], [1], True])
            entry.extend([[attacker['ship'], attacker.get('dmg', 0)]])

            if index != 0 and not entry[2]:
                entry[1][0] += 1
                entry[2] = True

    # 统计致命一击
    for attacker in kb_data.get('atts', []):
        if attacker.get('blow', True):
            corp_str = str(attacker.get('corp', ''))
            char_str = str(attacker.get('char', ''))
            if corp_str in corp_ids and char_str in char_dir:
                char_dir[char_str][0][0] += 1
                break  # 一条击杀邮件只统计一次致命一击


def generate_final_data(char_id, ship_id, char_dir):
    """把原始累加数据整理成最终报表结构"""
    new_dir = {}
    for char_key, records in char_dir.items():
        total_dmg = 0
        ship_types = set()
        char_name = char_id.get(char_key, f"Unknown({char_key})")

        for record in records[3:]:
            ship_id_str = str(record[0])
            ship_types.add(ship_id.get(ship_id_str, f"Unknown({ship_id_str})"))
            total_dmg += record[1]

        new_dir[char_key] = {
            'name': char_name,
            'sorties': records[1][0],  # 出击次数
            'kills': records[0][0],    # 击杀量（致命一击）
            'damage': total_dmg,
            'ships': '、'.join(ship_types),
        }
    return new_dir


# ============================== Excel 样式 ==============================

def apply_header_style(ws):
    """设置 Excel 表头样式"""
    bold_font = Font(bold=True, size=12)
    center_alignment = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    for col in range(1, 7):
        cell = ws.cell(row=1, column=col)
        cell.font = bold_font
        cell.alignment = center_alignment
        cell.border = thin_border


def apply_data_style(ws):
    """设置 Excel 数据区样式"""
    center_alignment = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = center_alignment
            cell.border = thin_border


def adjust_column_width(ws):
    """调整 Excel 列宽"""
    column_widths = {'A': 8, 'B': 35, 'C': 10, 'D': 10, 'E': 12, 'F': 25}
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width


def apply_chart_styles(chart):
    """去掉图表网格线与图例"""
    chart.y_axis.majorGridlines = None
    chart.x_axis.majorGridlines = None
    chart.legend = None


# ================================= 路由 =================================

@app.route('/', methods=['GET', 'POST'])
def index():
    """主页：提交 BR 链接并生成战斗统计报告"""
    latest_report = None
    error = None
    br_valid = False

    if request.method == 'POST':
        br_links = request.form.get('br_links', '').strip()
        report_date = request.form.get('report_date', '')
        custom_name = request.form.get('custom_name', '').strip()

        # 校验 BR 链接格式
        br_links_list = [url.strip() for url in br_links.split('\n') if url.strip()]
        invalid_links = [url for url in br_links_list if not re.match(cfg.BR_LINK_PATTERN, url)]

        if invalid_links:
            error = f"Invalid links: {', '.join(invalid_links)}"
            log_event("ERROR", f"Invalid BR links: {error}")
            return render_template('index.html', error=error, br_links_text=br_links)

        if not report_date or not custom_name:
            error = "Date and name are required"
            log_event("ERROR", "Missing date or name")
            return render_template('index.html', error=error, br_links_text=br_links)

        if not br_links_list:
            error = "Please enter valid BR links"
            log_event("ERROR", "No BR links provided")
            return render_template('index.html', error=error, br_links_text=br_links)

        with open(data_path('br_links.txt'), 'w', encoding='utf-8') as f:
            f.write(br_links)

        try:
            corp_ids = load_ids('corp_id.txt')
            char_id = load_ids('char_id.txt')
            ship_id = load_ids('ship_id.txt')

            # TODO(通用版): 如果你的军团/联盟 ID 还没配置，这里会给出可读的提示而不是空报告
            if not corp_ids:
                raise ValueError(
                    "尚未配置任何军团 ID，请先到「军团管理」页面添加你要统计的军团（ID 与名称）"
                )

            # 逐个校验 BR 链接可用性
            br_valid = True
            for br_url in br_links_list:
                br_id = extract_br_id(br_url)
                response = requests.get(cfg.BR_API_BASE.format(br_id), timeout=cfg.REQUEST_TIMEOUT)
                if response.status_code != 200:
                    br_valid = False
                    error = f"Invalid BR link: {br_url}"
                    log_event("ERROR", f"Invalid BR link: {br_url}")
                    break

            if br_valid:
                char_dir = {}

                for index_i, br_url in enumerate(br_links_list):
                    br_id = extract_br_id(br_url)
                    response_br = requests.get(cfg.BR_API_BASE.format(br_id), timeout=cfg.REQUEST_TIMEOUT)

                    if response_br.status_code != 200:
                        error = f"Failed to access: {br_url}"
                        log_event("ERROR", f"BR access failed: {br_url}")
                        continue

                    data = response_br.json()

                    if index_i != 0:
                        for char in char_dir:
                            char_dir[char][2] = False

                    for r in data.get('relateds', []):
                        for km in r.get('kms', []):
                            process_killmail(km, corp_ids, char_id, char_dir, index_i)

                new_dir = generate_final_data(char_id, ship_id, char_dir)

                wb = Workbook()
                ws = wb.active
                ws.title = "战斗统计"

                sorted_data = sorted(new_dir.items(), key=lambda x: x[1]['damage'], reverse=True)
                headers = ["排名", "角色", "出击次数", "击杀数", "总伤害", "使用舰船"]
                for idx, (_, item) in enumerate(sorted_data, start=1):
                    item['rank'] = idx

                ws.append(headers)
                apply_header_style(ws)

                for _, item in sorted_data:
                    ws.append([
                        item['rank'],
                        item['name'],
                        item['sorties'],
                        item['kills'],
                        item['damage'],
                        item['ships'],
                    ])

                apply_data_style(ws)
                adjust_column_width(ws)

                # 伤害图表
                damage_chart = BarChart()
                damage_chart.type = "col"
                damage_chart.title = f"总伤害排名前{cfg.CHART_TOP_N}"
                damage_data = Reference(ws, min_col=5, min_row=2, max_col=5,
                                        max_row=min(len(sorted_data), cfg.CHART_TOP_N + 1))
                damage_cats = Reference(ws, min_col=2, min_row=2,
                                        max_row=min(len(sorted_data), cfg.CHART_TOP_N + 1))
                damage_chart.add_data(damage_data, titles_from_data=False)
                damage_chart.set_categories(damage_cats)
                ws.add_chart(damage_chart, "H2")

                # 击杀数图表
                kills_chart = BarChart()
                kills_chart.type = "col"
                kills_chart.title = f"击杀数排名前{cfg.CHART_TOP_N}"
                kills_data = Reference(ws, min_col=4, min_row=2, max_col=4,
                                       max_row=min(len(sorted_data), cfg.CHART_TOP_N + 1))
                kills_cats = Reference(ws, min_col=2, min_row=2,
                                       max_row=min(len(sorted_data), cfg.CHART_TOP_N + 1))
                kills_chart.add_data(kills_data, titles_from_data=False)
                kills_chart.set_categories(kills_cats)
                ws.add_chart(kills_chart, "H18")

                filename = '{}_{}_{}.xlsx'.format(
                    cfg.REPORT_FILE_PREFIX,
                    report_date.replace("-", ""),
                    re.sub(r"[^\w-]", "_", custom_name),
                )
                save_path = data_path('reports', filename)
                os.makedirs(os.path.dirname(save_path), exist_ok=True)
                wb.save(save_path)

                latest_report = filename
                log_event("SUCCESS", f"Report generated: {filename}")

        except Exception as e:
            error = f"Report generation failed: {str(e)}"
            log_event("ERROR", f"Generation error: {str(e)}")

    # 取最新一份报告
    reports_dir = data_path('reports')
    if os.path.exists(reports_dir):
        reports = [f for f in os.listdir(reports_dir) if f.endswith('.xlsx')]
        if reports:
            latest_report = max(
                reports, key=lambda f: os.path.getctime(os.path.join(reports_dir, f))
            )

    return render_template(
        'index.html',
        latest_report=latest_report,
        error=error,
        br_valid=br_valid,
        br_links_text=request.form.get('br_links', '') or '\n'.join(load_br_links()),
    )


@app.route('/download-report/<path:filename>')
def download_report(filename):
    """下载报告（仅允许 data/reports 下已存在的 .xlsx）"""
    safe_path = resolve_report_path(filename)
    if not safe_path:
        log_event("ERROR", f"非法的报告下载请求: {filename}")
        return redirect(url_for('index'))
    try:
        return send_file(
            safe_path,
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    except Exception as e:
        log_event("ERROR", f"File download failed: {str(e)}")
        return redirect(url_for('index'))


@app.route('/delete-report/<path:filename>')
def delete_report(filename):
    """删除报告（仅允许 data/reports 下已存在的 .xlsx）"""
    safe_path = resolve_report_path(filename)
    if not safe_path:
        log_event("ERROR", f"非法的报告删除请求: {filename}")
        return jsonify(success=False, message="文件名非法"), 400
    try:
        os.remove(safe_path)
        log_event("INFO", f"Deleted report: {filename}")
        return jsonify(success=True)
    except Exception as e:
        log_event("ERROR", f"Failed to delete report: {str(e)}")
        return jsonify(success=False)


# TODO(通用版): 角色管理目前只用于补充/修正角色 ID 与名称映射，可按需扩展白名单等功能
@app.route('/manage-chars', methods=['GET', 'POST'])
def manage_chars():
    """管理角色 ID 与名称的映射（data/char_id.txt）"""
    char_id = load_ids('char_id.txt')
    ship_id = load_ids('ship_id.txt')

    if request.method == 'POST':
        if 'add_char' in request.form:
            new_id = request.form['char_id'].strip()
            new_name = request.form['char_name'].strip()
            ship = request.form.get('ship')

            if new_id and new_name:
                char_id[new_id] = new_name
                save_ids('char_id.txt', char_id)
                if ship and ship not in ship_id.values():
                    new_ship_id = str(max([int(k) for k in ship_id.keys()] or [0]) + 1)
                    ship_id[new_ship_id] = ship
                    save_ids('ship_id.txt', ship_id)
                log_event("INFO", f"Added character: {new_name}(ID: {new_id})")
        elif 'delete' in request.form:
            del_id = request.form['delete']
            if del_id in char_id:
                del char_id[del_id]
                save_ids('char_id.txt', char_id)
                log_event("INFO", f"Deleted character ID: {del_id}")
        return redirect(url_for('manage_chars'))

    return render_template('manage_chars.html', chars=char_id, ships=sorted(ship_id.values()))


@app.route('/manage-corps', methods=['GET', 'POST'])
def manage_corps():
    """管理要统计的军团 ID 与名称（data/corp_id.txt）—— 通用版核心配置入口"""
    corps = load_ids('corp_id.txt')

    if request.method == 'POST':
        if 'add_corp' in request.form:
            corp_id = request.form['corp_id'].strip()
            corp_name = request.form['corp_name'].strip()
            if corp_id and corp_name:
                corps[corp_id] = corp_name
                save_ids('corp_id.txt', corps)
                log_event("INFO", f"Added corporation: {corp_name}(ID: {corp_id})")
        elif 'delete' in request.form:
            corp_id = request.form['delete']
            if corp_id in corps:
                del corps[corp_id]
                save_ids('corp_id.txt', corps)
                log_event("INFO", f"Deleted corporation ID: {corp_id}")
        return redirect(url_for('manage_corps'))

    return render_template('manage_corps.html', corps=corps)


@app.route('/history')
def history():
    """历史报告列表"""
    reports_dir = data_path('reports')
    reports = []
    if os.path.exists(reports_dir):
        reports = sorted(
            [{'name': f,
              'ctime': datetime.fromtimestamp(os.path.getctime(os.path.join(reports_dir, f)))}
             for f in os.listdir(reports_dir) if f.endswith('.xlsx')],
            key=lambda x: x['ctime'],
            reverse=True,
        )
    return render_template('history.html', reports=reports)


@app.route('/logs')
def logs():
    """系统日志"""
    logs_content = []
    try:
        with open(data_path('system.log'), 'r', encoding='utf-8') as f:
            logs_content = f.readlines()
    except FileNotFoundError:
        pass
    return render_template('logs.html', logs=logs_content)


@app.route('/docs')
def docs():
    """内置操作文档"""
    return render_template('docs.html')


def run_app():
    """启动 Web 服务"""
    os.makedirs(data_path('reports'), exist_ok=True)

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(data_path('detailed.log'), encoding='utf-8'),
            logging.StreamHandler(),
        ],
    )

    logging.info(f"Starting {cfg.SITE_NAME} v{cfg.SITE_VERSION} ...")
    logging.info(f"Application data folder: {app.config['DATA_FOLDER']}")
    logging.info(f"Web interface: http://{cfg.HOST}:{cfg.PORT}")
    logging.info("Close this window to stop the application\n")

    if cfg.AUTO_OPEN_BROWSER:
        import webbrowser
        webbrowser.open(f'http://{cfg.HOST}:{cfg.PORT}')

    app.run(host=cfg.HOST, port=cfg.PORT)


if __name__ == '__main__':
    run_app()
