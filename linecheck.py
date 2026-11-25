# -*- coding: utf-8 -*-
"""线条闭合检测的纯几何部分（不依赖 CorelDRAW，方便单测）。

坐标系：一律用**毫米**。调用方负责把文档单位换算过来。

两类检测：
  boundary  边界检测：只看「边界」图层的线条，找贴近但没接上的端点
  global    全图闭合：参考集是全部线条（含闭合线），找悬空端点
判定：端点 E 到参考线段的最近距离 d
      0 < eps < d <= 阈值  → 未闭合（有缝）
      d <= eps             → 已接触，不算问题
      d >  阈值            → 离得远，不算问题
延长：以「端点 E − 前一个节点 P」为方向作射线，与参考线段求最近交点，
      得到延长终点 I（供用户确认后写入图形）。
"""
import math

EPS_MM = 0.002          # 2 微米以内视为已接触
MAX_EXTEND_MM = 60.0    # 延长超过 60mm 视为异常，不给方案


# ---------------------------------------------------------------- 基础几何

def pt_seg_dist(p, a, b):
    ax, ay = a
    bx, by = b
    px, py = p
    dx, dy = bx - ax, by - ay
    l2 = dx * dx + dy * dy
    if l2 <= 1e-18:
        return math.hypot(px - ax, py - ay)
    t = ((px - ax) * dx + (py - ay) * dy) / l2
    t = 0.0 if t < 0 else (1.0 if t > 1 else t)
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


def ray_seg_hit(p, d, a, b):
    """射线 p + t·d 与线段 ab 求交。返回 t（>0）或 None。"""
    rx, ry = d
    sx, sy = b[0] - a[0], b[1] - a[1]
    den = rx * sy - ry * sx
    if abs(den) < 1e-15:
        return None
    qx, qy = a[0] - p[0], a[1] - p[1]
    t = (qx * sy - qy * sx) / den
    u = (qx * ry - qy * rx) / den
    if t > 1e-9 and -1e-9 <= u <= 1 + 1e-9:
        return t
    return None


def segments_of(pts, closed):
    segs = []
    for i in range(len(pts) - 1):
        segs.append((pts[i], pts[i + 1]))
    if closed and len(pts) > 2:
        segs.append((pts[-1], pts[0]))
    return segs


def polyline_length(pts):
    n = 0.0
    for i in range(len(pts) - 1):
        n += math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1])
    return n


# ---------------------------------------------------------------- 主检测

def analyze(geo, mode, thr_mm, boundary_layers=("边界",), eps_mm=EPS_MM,
            max_extend_mm=MAX_EXTEND_MM):
    """geo: {图层名: [{"addr": [...], "closed": bool, "pts": [(x_mm,y_mm), ...]}]}

    返回 issues 列表，每项：
      layer, addr, sub, end(0=首端 1=尾端), pt, prev, dist_mm,
      ref_layer, ref_addr, extend_to (x,y) 或 None, extend_len_mm
    """
    # 参考集：边界模式只取用户选中的图层（可多选）；全图模式取全部
    ref_sel = set(boundary_layers or ())
    refs = []
    for lname, paths in geo.items():
        if mode == "boundary":
            if lname not in ref_sel:
                continue
        for idx, sp in enumerate(paths):
            for seg in segments_of(sp["pts"], sp["closed"]):
                refs.append((lname, tuple(sp["addr"]), idx, seg[0], seg[1]))

    issues = []
    for lname, paths in geo.items():
        if mode == "boundary" and lname in ref_sel:
            continue                       # 参考图层自身不作为待检对象
        for idx, sp in enumerate(paths):
            if sp["closed"]:
                continue                   # 闭合线不检（用户要求）
            pts = sp["pts"]
            if len(pts) < 2:
                continue
            for end in (0, 1):
                e = pts[0] if end == 0 else pts[-1]
                p = pts[1] if end == 0 else pts[-2]
                best = None
                for (rl, raddr, ridx, a, b) in refs:
                    # 同一对象的子路径不当参考（避免自己贴自己）
                    if rl == lname and raddr == tuple(sp["addr"]) and ridx == idx:
                        continue
                    d = pt_seg_dist(e, a, b)
                    if best is None or d < best[0]:
                        best = (d, rl, raddr, ridx, a, b)
                if best is None:
                    continue
                d = best[0]
                if d <= eps_mm or d > thr_mm:
                    continue
                # 延长方向：前一个节点 → 端点
                dx, dy = e[0] - p[0], e[1] - p[1]
                n = math.hypot(dx, dy)
                if n < 1e-9:
                    dx, dy, n = 1.0, 0.0, 1.0
                dx, dy = dx / n, dy / n
                hit = None
                for (rl, raddr, ridx, a, b) in refs:
                    if rl == lname and raddr == tuple(sp["addr"]) and ridx == idx:
                        continue
                    t = ray_seg_hit(e, (dx, dy), a, b)
                    if t is None or t * n > max_extend_mm:
                        continue
                    if hit is None or t < hit[0]:
                        hit = (t, rl, a, b)
                extend_to = None
                extend_len = None
                ref_layer = best[1]
                ref_addr = best[2]
                if hit is not None:
                    extend_to = (e[0] + dx * hit[0], e[1] + dy * hit[0])
                    extend_len = hit[0]
                    ref_layer = hit[1]
                issues.append({
                    "layer": lname,
                    "addr": list(sp["addr"]),
                    "sub": idx,
                    "end": end,
                    "pt": (round(e[0], 4), round(e[1], 4)),
                    "prev": (round(p[0], 4), round(p[1], 4)),
                    "dist_mm": round(d, 4),
                    "ref_layer": ref_layer,
                    "ref_addr": list(ref_addr),
                    "extend_to": (round(extend_to[0], 4), round(extend_to[1], 4)) if extend_to else None,
                    "extend_len_mm": round(extend_len, 3) if extend_len else None,
                })
    issues.sort(key=lambda x: (x["layer"], x["dist_mm"]))
    return issues
