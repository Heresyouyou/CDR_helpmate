# -*- coding: utf-8 -*-
"""
自检：验证 CDR 助手与 CorelDRAW 的完整链路。

会临时改动当前文档的「默认对象属性」并读回验证，最后按原值恢复。
建议先存盘再跑。

用法：  "D:\\Python 3.12.4\\python.exe" selfcheck.py
"""

import importlib.util
import os
import sys

import pythoncom


HERE = os.path.dirname(os.path.abspath(__file__))


def load_tool():
    spec = importlib.util.spec_from_file_location(
        "cdtool", os.path.join(HERE, "cdr_defaults.pyw"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main():
    cd = load_tool()
    pythoncom.CoInitialize()

    bridge = cd.CdrBridge()
    if cd.CdrBridge.cdr_running() is False:
        print("[跳过] CorelDRAW 未运行，请先打开并新建/打开一个文档")
        return 1

    ok, msg = bridge.attach()
    print("1) 连接        :", msg)
    if not ok or bridge.doc is None:
        print("[失败] 需要一个打开的文档")
        return 1

    print("2) 单位常量    :", bridge.resolve_mm_unit())
    orig = bridge.read_current()
    print("3) 当前默认值  :", orig)
    if not orig:
        print("[失败] 读不到默认样式")
        return 1

    test = {
        "width_mm": 0.5,
        "outline_rgb": (200, 20, 40),
        "fill_mode": "none",
        "fill_rgb": (210, 220, 230),
        "kinds": {"graphic": True, "artistic": True},
        "line_caps": 2,
        "line_join": 1,
        "justify": 2,
        "scale_with": True,
        "behind_fill": False,
    }
    print("4) 写入测试值  :", bridge.apply(test))

    layer = bridge.doc.ActiveLayer
    probe = layer.CreateRectangle(0.5, 0.5, 2.0, 2.0)
    got_w = bridge.from_doc(probe.Outline.Width)
    got_c = probe.Outline.Color.HexValue
    got_cap = probe.Outline.LineCaps
    got_join = probe.Outline.LineJoin
    got_just = probe.Outline.Justification
    got_scale = probe.Outline.ScaleWithShape
    probe.Delete()

    print("5) 新对象继承  : %.4f mm | %s | 帽=%r 角=%r 位=%r 随缩放=%r"
          % (got_w, got_c, got_cap, got_join, got_just, got_scale))

    checks = [
        ("线宽", abs(got_w - 0.5) < 0.02),
        ("颜色", str(got_c).upper() == "#C81428"),
        ("线帽", got_cap == 2),
        ("线角", got_join == 1),
        ("轮廓位置", got_just == 2),
        ("随对象缩放", bool(got_scale) is True),
    ]

    back = bridge.read_current()
    print("6) 回读        :", back)
    checks.append(("回读线宽", back.get("width_mm") == 0.5))
    checks.append(("回读线帽", back.get("line_caps") == 2))

    print("7) 填充写入    :", bridge.apply(dict(test, fill_mode="custom")))
    probe = layer.CreateRectangle(0.5, 0.5, 2.0, 2.0)
    fill_ok = int(probe.Fill.Type) == 1
    probe.Delete()
    checks.append(("填充", fill_ok))

    print("8) 发丝线      :", bridge.apply(dict(test, width_mm=0.0)))
    checks.append(("发丝线", bridge.read_current().get("width_mm") == 0.0))

    print("9) 线型（需先选中对象）")
    probe = layer.CreateRectangle(0.5, 0.5, 2.0, 2.0)
    try:
        probe.CreateSelection()
        n, m = bridge.apply_linestyle(4)
        print("   ", n, m)
        checks.append(("线型作用于选区", n == 1))
    except Exception as exc:
        print("    失败:", exc)
        checks.append(("线型作用于选区", False))
    probe.Delete()

    restore = dict(test)
    restore.update({
        "width_mm": orig.get("width_mm", 0.2),
        "outline_rgb": orig.get("outline_rgb") or (0, 0, 0),
        "fill_mode": orig.get("fill_mode") or "none",
        "fill_rgb": orig.get("fill_rgb") or (217, 217, 218),
        "line_caps": orig.get("line_caps", 0),
        "line_join": orig.get("line_join", 0),
        "justify": orig.get("justify", 0),
        "scale_with": orig.get("scale_with", False),
        "behind_fill": orig.get("behind_fill", False),
    })
    print("10) 按原值恢复 :", bridge.apply(restore))
    print("    恢复后回读 :", bridge.read_current())

    print("\n结果汇总")
    bad = 0
    for name, good in checks:
        print("  [%s] %s" % ("PASS" if good else "FAIL", name))
        if not good:
            bad += 1
    print("\n%s" % ("全部通过" if bad == 0 else "有 %d 项未通过" % bad))
    return 0 if bad == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
