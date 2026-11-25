# -*- coding: utf-8 -*-
"""生成应用图标 app.ico（深色圆角 + 一条样条曲线 + 两个控制点）。"""

import math
import os

from PIL import Image, ImageDraw

S = 256
BG = (47, 62, 70, 255)
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

d.rounded_rectangle([10, 10, S - 10, S - 10], radius=54, fill=BG)

pts = []
for i in range(0, 121):
    t = i / 120.0
    x = 42 + t * (S - 84)
    y = S / 2.0 + math.sin(t * math.pi * 2) * 40
    pts.append((x, y))

d.line(pts, fill=(255, 255, 255, 255), width=15, joint="curve")
r = 11
d.ellipse([pts[0][0] - r, pts[0][1] - r, pts[0][0] + r, pts[0][1] + r], fill=(230, 40, 50, 255))
d.ellipse([pts[-1][0] - r, pts[-1][1] - r, pts[-1][0] + r, pts[-1][1] + r], fill=(60, 170, 220, 255))

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.ico")
img.save(out, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("icon:", out, os.path.getsize(out), "bytes")
