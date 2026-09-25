# -*- coding: utf-8 -*-
"""
出勤统计蓝图
--------------------------------------------------------------------------
通用版改动说明：
  1. 出勤数据从「内存列表」改为落盘到 data/attendance.json，重启程序不丢数据；
  2. 导出 Excel 改为内存流（BytesIO）输出，不再往工作目录写 output.xlsx；
  3. 活动类型选项来自 config.EVENT_TYPES，可按自己的玩法自行增删。
"""

import io
import json
import os
from collections import defaultdict
from datetime import datetime

import pandas as pd
from flask import Blueprint, current_app, jsonify, render_template, request, send_file

import config as cfg

attendance_bp = Blueprint('attendance', __name__)

# 出勤数据文件名（存放在 config.get_data_folder() 指定的数据目录下）
STORE_FILENAME = "attendance.json"


# ============================== 数据持久化 ==============================

def _store_path():
    """出勤数据文件的绝对路径"""
    return os.path.join(current_app.config['DATA_FOLDER'], STORE_FILENAME)


def _load_submissions():
    """读取全部出勤记录，文件不存在或损坏时返回空列表"""
    path = _store_path()
    if not os.path.exists(path):
        return []
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception as e:
        current_app.logger.error(f"读取出勤数据失败: {e}")
        return []


def _save_submissions(submissions):
    """写回全部出勤记录"""
    path = _store_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(submissions, f, ensure_ascii=False, indent=2)


def _build_stats(submissions):
    """按角色名汇总出勤次数，返回已排序的列表"""
    stats = defaultdict(int)
    for sub in submissions:
        for member in sub.get('members', []):
            stats[member] += 1
    stats_list = [{'name': k, 'count': v} for k, v in stats.items()]
    stats_list.sort(key=lambda x: (-x['count'], x['name']))
    return stats_list


# ================================ 接口 ================================

@attendance_bp.route('/submit_fleet', methods=['POST'])
def submit_fleet():
    """提交一次舰队出勤的角色名单"""
    try:
        fleet_date = request.form['fleet_date']
        event_type = request.form['event_type']
        fleet_note = request.form.get('fleet_note', '')
        # 去重且保序，避免同一份名单里重复角色被重复计数
        members = list(dict.fromkeys(
            m.strip() for m in request.form['fleet_members'].split('\n') if m.strip()
        ))

        if not members:
            raise ValueError("角色列表不能为空")

        submissions = _load_submissions()
        submissions.append({
            'timestamp': datetime.now().isoformat(),
            'date': fleet_date,
            'event_type': event_type,
            'note': fleet_note,
            'members': members,
            'member_count': len(members),
        })
        _save_submissions(submissions)

        return jsonify({
            'status': 'success',
            'count': len(members),
            'message': f'成功添加 {len(members)} 个角色',
        })

    except Exception as e:
        current_app.logger.error(f"提交失败: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 400


@attendance_bp.route('/get_stats')
def get_stats():
    """返回角色出勤次数排行"""
    try:
        return jsonify(_build_stats(_load_submissions()))
    except Exception as e:
        current_app.logger.error(f"获取统计失败: {e}")
        return jsonify([])


@attendance_bp.route('/get_history')
def get_history():
    """返回简化的提交历史，按时间倒序"""
    try:
        history = []
        for sub in sorted(_load_submissions(), key=lambda x: x['timestamp'], reverse=True):
            history.append({
                'timestamp': sub['timestamp'],
                'date': sub['date'],
                'event_type': sub['event_type'],
                'member_count': sub['member_count'],
                'note': sub['note'],
            })
        return jsonify(history)
    except Exception as e:
        current_app.logger.error(f"获取历史失败: {e}")
        return jsonify([])


@attendance_bp.route('/export_stats')
def export_stats():
    """导出出勤统计 Excel（统计表 + 历史记录两个工作表）"""
    try:
        submissions = _load_submissions()
        if not submissions:
            raise ValueError("没有数据可导出")

        stats_df = pd.DataFrame(
            [(item['name'], item['count']) for item in _build_stats(submissions)],
            columns=['角色名', '出现次数'],
        )

        history_df = pd.DataFrame([{
            '提交时间': datetime.fromisoformat(sub['timestamp']).strftime('%Y-%m-%d %H:%M:%S'),
            '活动日期': sub['date'],
            '活动类型': sub['event_type'],
            '角色数量': sub['member_count'],
            '活动说明': sub['note'],
            '角色列表': ', '.join(sub['members']),
        } for sub in submissions])

        # 直接写入内存流，不产生临时文件
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            stats_df.to_excel(writer, sheet_name='出勤统计', index=False)
            history_df.to_excel(writer, sheet_name='历史记录', index=False)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name=f'舰队出勤统计_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    except Exception as e:
        current_app.logger.error(f"导出失败: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500


@attendance_bp.route('/has_data')
def has_data():
    return jsonify({'has_data': len(_load_submissions()) > 0})


@attendance_bp.route('/attendance')
def attendance_index():
    return render_template('attendance.html', event_types=cfg.EVENT_TYPES)
