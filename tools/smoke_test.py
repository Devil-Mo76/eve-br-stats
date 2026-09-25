# -*- coding: utf-8 -*-
"""
冒烟测试脚本
--------------------------------------------------------------------------------
用途：不联网、不动真实数据，快速验证页面与接口是否正常。
原理：把数据目录临时指向系统临时目录（通过环境变量 BR_DATA_DIR），
      因此不会污染 data/ 下的真实数据文件。

用法：
    python tools/smoke_test.py
"""

import os
import shutil
import sys
import tempfile

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

# ---- 关键：在导入 app 之前把数据目录指到临时目录 ----
TMP_DATA = tempfile.mkdtemp(prefix="br_smoke_")
os.environ["BR_DATA_DIR"] = TMP_DATA

# 复制一份舰船数据过去，保证报表逻辑可用
src_ship = os.path.join(PROJECT_ROOT, "data", "ship_id.txt")
if os.path.exists(src_ship):
    shutil.copy(src_ship, os.path.join(TMP_DATA, "ship_id.txt"))

import app as app_module  # noqa: E402

PASS, FAIL = [], []


def check(name, condition, extra=""):
    (PASS if condition else FAIL).append(name)
    print(f"{'[PASS]' if condition else '[FAIL]'} {name}{(' -> ' + extra) if extra else ''}")


def main():
    app_module.app.config['TESTING'] = True
    client = app_module.app.test_client()
    print(f"临时数据目录: {TMP_DATA}\n")

    # ---------- 1. 页面渲染 ----------
    for path in ['/', '/docs', '/manage-corps', '/manage-chars', '/history', '/logs']:
        r = client.get(path)
        check(f"GET {path} 返回 200", r.status_code == 200, f"实际 {r.status_code}")

    r = client.get('/attendance/attendance')
    check("GET /attendance/attendance 返回 200（出勤页真实地址）", r.status_code == 200,
          f"实际 {r.status_code}")

    r = client.get('/attendance')
    check("GET /attendance 返回 404（蓝图前缀叠加，属预期行为）", r.status_code == 404,
          f"实际 {r.status_code}")

    # ---------- 2. 品牌信息已通用化 ----------
    html = client.get('/').get_data(as_text=True)
    check("首页不再出现 Cosmic-Wanderers", "Cosmic-Wanderers" not in html)
    check("首页使用了可配置系统名", "EVE Online BR 战斗统计系统" in html)
    check("导航含「管理」下拉入口", "军团管理" in html and "角色管理" in html)

    # ---------- 3. 军团管理写入 ----------
    r = client.post('/manage-corps', data={'add_corp': '1', 'corp_id': '12345678',
                                           'corp_name': 'SmokeTest Corp'},
                    follow_redirects=True)
    check("POST /manage-corps 添加军团成功", r.status_code == 200)
    corp_file = os.path.join(TMP_DATA, 'corp_id.txt')
    content = open(corp_file, encoding='utf-8').read() if os.path.exists(corp_file) else ''
    check("corp_id.txt 已写入军团", '12345678_SmokeTest Corp' in content, content.strip())

    # ---------- 4. 出勤统计 ----------
    members = "Pilot Alpha\nPilot Beta\nPilot Alpha\n"  # 含重复，用于验证去重
    r = client.post('/attendance/submit_fleet', data={
        'fleet_date': '2025-09-25', 'event_type': '小队活动',
        'fleet_note': 'smoke', 'fleet_members': members,
    })
    check("POST /attendance/submit_fleet 提交成功", r.status_code == 200,
          r.get_data(as_text=True)[:120])
    check("提交返回的角色数已去重（应为 2）", r.get_json().get('count') == 2,
          str(r.get_json()))

    stats = client.get('/attendance/get_stats').get_json()
    check("GET /attendance/get_stats 返回 2 个角色", len(stats) == 2, str(stats))
    check("每名角色出勤次数为 1", all(s['count'] == 1 for s in stats), str(stats))

    check("出勤数据已落盘 attendance.json",
          os.path.exists(os.path.join(TMP_DATA, 'attendance.json')))

    r = client.get('/attendance/export_stats')
    check("GET /attendance/export_stats 导出成功", r.status_code == 200,
          f"实际 {r.status_code}")
    check("导出内容为 xlsx", r.data[:2] == b'PK', str(r.data[:4]))

    # ---------- 5. 安全检查 ----------
    r = client.get('/download-report/..%2F..%2F..%2Fwindows%2Fwin.ini')
    check("目录穿越下载被拦截", r.status_code in (302, 400, 404), f"实际 {r.status_code}")

    r = client.get('/delete-report/..%2F..%2Fcorp_id.txt')
    check("目录穿越删除被拦截", r.status_code in (400, 404), f"实际 {r.status_code}")
    check("corp_id.txt 未被越权删除", os.path.exists(corp_file))

    # ---------- 6. 未配置军团时的提示 ----------
    os.remove(corp_file)
    r = client.post('/', data={'br_links': 'https://br.evetools.org/br/abc123',
                               'report_date': '2025-09-25', 'custom_name': 'smoke'})
    body = r.get_data(as_text=True)
    check("未配置军团时给出明确提示", '尚未配置任何军团' in body)

    r = client.post('/', data={'br_links': 'not-a-link',
                               'report_date': '2025-09-25', 'custom_name': 'smoke'})
    check("非法 BR 链接被拒绝",
          'Invalid links' in r.get_data(as_text=True))

    # ---------- 汇总 ----------
    print(f"\n===== 结果：通过 {len(PASS)} 项，失败 {len(FAIL)} 项 =====")
    if FAIL:
        print("失败项：")
        for f in FAIL:
            print(f"  - {f}")
    return 1 if FAIL else 0


if __name__ == '__main__':
    code = main()
    shutil.rmtree(TMP_DATA, ignore_errors=True)
    sys.exit(code)
