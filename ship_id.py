# -*- coding: utf-8 -*-
"""
舰船数据更新脚本
--------------------------------------------------------------------------------
作用：从 EVE ESI 接口拉取当前全部已发布舰船的 类型ID → 名称 映射，
      写入 data/ship_id.txt，供战斗统计报表把「舰船ID」翻译成可读的舰船名。

什么时候需要跑：
    报表里出现 Unknown(数字ID) 形式的舰船，说明新版本加了新船，
    此时执行本脚本，生成后覆盖 data/ship_id.txt 即可。

用法：
    python ship_id.py                     # 直接覆盖 data/ship_id.txt
    python ship_id.py --output other.txt  # 输出到指定文件
"""

import argparse
import os
import time

import requests

import config as cfg


def get_eve_online_ships():
    """拉取全部已发布舰船的 {type_id: name}"""
    category_id = 6  # 舰船类别 ID
    lang = "en"      # 语言：en / zh 等，按需修改

    category_url = f"{cfg.ESI_BASE}/universe/categories/{category_id}/"
    response = requests.get(category_url, timeout=cfg.REQUEST_TIMEOUT)
    if response.status_code != 200:
        print(f"获取舰船分类失败，状态码: {response.status_code}")
        return {}

    category_data = response.json()
    if 'groups' not in category_data:
        print(f"响应中缺少 groups 字段: {category_data}")
        return {}

    ships = {}
    for group_id in category_data['groups']:
        group_url = f"{cfg.ESI_BASE}/universe/groups/{group_id}/?language={lang}"
        response = requests.get(group_url, timeout=cfg.REQUEST_TIMEOUT)
        if response.status_code == 200:
            group_data = response.json()
            if group_data.get('published', False):  # 只取已发布的舰船分组
                for type_id in group_data['types']:
                    ship_url = f"{cfg.ESI_BASE}/universe/types/{type_id}/?language={lang}"
                    resp = requests.get(ship_url, timeout=cfg.REQUEST_TIMEOUT)
                    if resp.status_code == 200:
                        ship_data = resp.json()
                        if ship_data.get('published', False):
                            ships[type_id] = ship_data['name']
                            print(f"已收录: {ship_data['name']} (ID: {type_id})")
                    else:
                        print(f"获取类型 {type_id} 失败，状态码: {resp.status_code}")
                    time.sleep(0.1)  # 限速，避免触发接口速率限制
        else:
            print(f"获取分组 {group_id} 失败，状态码: {response.status_code}")
        time.sleep(0.1)

    return ships


def save_ships_to_file(ships, file_path):
    """按「ID_名称」每行一条写入文件，与 data/ship_id.txt 格式保持一致"""
    os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)
    with open(file_path, "w", encoding="utf-8") as f:
        for ship_id, ship_name in ships.items():
            f.write(f"{ship_id}_{ship_name}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="更新 EVE 舰船 ID 与名称映射表")
    parser.add_argument(
        "--output", "-o",
        default=os.path.join(cfg.get_data_folder(), "ship_id.txt"),
        help="输出文件路径，默认写入数据目录下的 ship_id.txt",
    )
    args = parser.parse_args()

    ship_dict = get_eve_online_ships()
    save_ships_to_file(ship_dict, args.output)
    print(f"\n舰船数据已保存到: {args.output}")
    print(f"共获取舰船数量: {len(ship_dict)}")
