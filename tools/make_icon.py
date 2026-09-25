# -*- coding: utf-8 -*-
"""
图标生成脚本（通用版自带）
--------------------------------------------------------------------------------
用途：生成两套图标
    1. static/favicon.ico   浏览器标签页图标
    2. app_icon.ico         PyInstaller 打包 EXE 用的图标
两套图标使用同一套几何设计（深色圆角底 + 青色准星 + 统计柱），
你可以直接改下面的 COLORS / 尺寸参数，或把本文件换成自己的图片。

用法：
    python tools/make_icon.py
依赖：
    pip install pillow
"""

import os
import sys

from PIL import Image, ImageDraw

# TODO(通用版): 想换配色就改这两行
BG_OUTER = (22, 36, 58, 255)      # 背板渐变的起始色（左上）
BG_INNER = (5, 8, 15, 255)        # 背板渐变的结束色（右下）
ACCENT = (0, 191, 255, 255)       # 主强调色（准星 / 描边）
BAR_TOP = (0, 208, 255, 255)      # 柱状图顶部颜色
BAR_BOTTOM = (0, 119, 194, 255)   # 柱状图底部颜色

SIZE = 256
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ICO_SIZES = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


def vertical_gradient(size, top_color, bottom_color):
    """生成一张竖向渐变图，用于柱子的填色"""
    grad = Image.new("RGBA", (1, size), (0, 0, 0, 0))
    for y in range(size):
        ratio = y / max(size - 1, 1)
        grad.putpixel((0, y), tuple(
            int(top_color[i] + (bottom_color[i] - top_color[i]) * ratio) for i in range(4)
        ))
    return grad.resize((size, size))


def diagonal_gradient(size, start_color, end_color):
    """生成一张对角渐变图，用于背板"""
    grad = Image.new("RGBA", (size, size))
    for y in range(size):
        for x in range(size):
            ratio = (x + y) / (2 * (size - 1))
            grad.putpixel((x, y), tuple(
                int(start_color[i] + (end_color[i] - start_color[i]) * ratio) for i in range(4)
            ))
    return grad


def build_icon(size=SIZE):
    """按几何设计绘制一枚图标"""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    scale = size / 256

    # 背板：圆角矩形 + 对角渐变
    board_mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(board_mask).rounded_rectangle(
        [8 * scale, 8 * scale, 248 * scale, 248 * scale],
        radius=48 * scale, fill=255,
    )
    img.paste(diagonal_gradient(size, BG_OUTER, BG_INNER), (0, 0), board_mask)

    draw = ImageDraw.Draw(img)

    # 背板描边
    draw.rounded_rectangle(
        [8 * scale, 8 * scale, 248 * scale, 248 * scale],
        radius=48 * scale, outline=ACCENT[:3] + (115,), width=max(1, int(4 * scale)),
    )

    # 准星：外环 + 内环
    def ring(cx, cy, r, width, alpha=255):
        draw.ellipse(
            [(cx - r) * scale, (cy - r) * scale, (cx + r) * scale, (cy + r) * scale],
            outline=ACCENT[:3] + (alpha,), width=max(1, int(width * scale)),
        )

    ring(128, 112, 60, 6, 140)
    ring(128, 112, 22, 8)

    # 准星四向刻度
    tick = max(1, int(8 * scale))
    for x1, y1, x2, y2 in [(128, 34, 128, 58), (128, 166, 128, 190),
                           (50, 112, 74, 112), (182, 112, 206, 112)]:
        draw.line([x1 * scale, y1 * scale, x2 * scale, y2 * scale],
                  fill=ACCENT, width=tick)

    # 底部统计柱：用竖向渐变贴上去，模拟高度差
    bars = [(70, 196, 24, 30), (104, 184, 24, 42),
            (138, 170, 24, 56), (172, 158, 24, 68)]
    for bx, by, bw, bh in bars:
        box = [bx * scale, by * scale, (bx + bw) * scale, (by + bh) * scale]
        mask = Image.new("L", (size, size), 0)
        ImageDraw.Draw(mask).rounded_rectangle(box, radius=6 * scale, fill=255)
        grad = vertical_gradient(int(bh * scale), BAR_TOP, BAR_BOTTOM)
        layer = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        layer.paste(grad, (int(bx * scale), int(by * scale)))
        img.paste(layer, (0, 0), mask)

    return img


def main():
    master = build_icon(SIZE)

    # 1. 浏览器 favicon
    favicon_path = os.path.join(PROJECT_ROOT, "static", "favicon.ico")
    master.save(favicon_path, format="ICO", sizes=ICO_SIZES)
    print(f"[OK] 已生成浏览器图标: {favicon_path}")

    # 2. PyInstaller EXE 图标
    exe_icon_path = os.path.join(PROJECT_ROOT, "app_icon.ico")
    master.save(exe_icon_path, format="ICO", sizes=ICO_SIZES)
    print(f"[OK] 已生成 EXE 图标: {exe_icon_path}")

    # 3. 顺手导出一张 PNG，方便当 LOGO 用
    png_path = os.path.join(PROJECT_ROOT, "static", "site_logo.png")
    master.save(png_path, format="PNG")
    print(f"[OK] 已导出 PNG: {png_path}")


if __name__ == "__main__":
    try:
        main()
    except ImportError:
        print("缺少依赖 Pillow，请先执行： pip install pillow")
        sys.exit(1)
