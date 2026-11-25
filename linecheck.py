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


def foot_on_seg(p, a, b):
    """点在直线段上的垂足（夹在段内）。"""
    ax, ay = a
    bx, by = b
    dx, dy = bx - ax, by - ay
    l2 = dx * dx + dy * dy
    if l2 <= 1e-18:
        return a
    t = ((p[0] - ax) * dx + (p[1] - ay) * dy) / l2
    t = 0.0 if t < 0 else (1.0 if t > 1 else t)
    return (ax + t * dx, ay + t * dy)


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


def _seg_dist_pair(a1, a2, b1, b2):
    """两线段最短距离；相交返回 0。"""
    d1 = (a2[0] - a1[0], a2[1] - a1[1])
    d2 = (b2[0] - b1[0], b2[1] - b1[1])
    den = d1[0] * d2[1] - d1[1] * d2[0]
    if abs(den) > 1e-15:
        t = ((b1[0] - a1[0]) * d2[1] - (b1[1] - a1[1]) * d2[0]) / den
        u = ((b1[0] - a1[0]) * d1[1] - (b1[1] - a1[1]) * d1[0]) / den
        if -1e-9 <= t <= 1 + 1e-9 and -1e-9 <= u <= 1 + 1e-9:
            return 0.0

    def seg_pt(p, s1, s2):
        vx, vy = s2[0] - s1[0], s2[1] - s1[1]
        l2 = vx * vx + vy * vy
        if l2 <= 1e-18:
            return math.hypot(p[0] - s1[0], p[1] - s1[1])
        tt = ((p[0] - s1[0]) * vx + (p[1] - s1[1]) * vy) / l2
        tt = 0.0 if tt < 0 else (1.0 if tt > 1 else tt)
        return math.hypot(p[0] - (s1[0] + tt * vx), p[1] - (s1[1] + tt * vy))

    return min(seg_pt(a1, b1, b2), seg_pt(a2, b1, b2),
               seg_pt(b1, a1, a2), seg_pt(b2, a1, a2))


def closest_pair(a1, a2, b1, b2):
    """两线段之间的最近点对 (p, q)。"""
    cands = [(a1, foot_on_seg(a1, b1, b2)), (a2, foot_on_seg(a2, b1, b2)),
             (foot_on_seg(b1, a1, a2), b1), (foot_on_seg(b2, a1, a2), b2)]
    best = None
    for p, q in cands:
        d = math.hypot(p[0] - q[0], p[1] - q[1])
        if best is None or d < best[0]:
            best = (d, p, q)
    return best[1], best[2], best[0]


def _already_joined(by_key, ka, kb, pt, radius):
    """两条子路径是否在 pt 附近（radius 内）已经有相交/相接。

    有的话，这个缝漏进去的只是一小片薄区（≤radius），肉眼看不出来，
    属于「修不修都一样」的低风险项。
    """
    sa = by_key.get(ka) or []
    sb = by_key.get(kb) or []
    r2 = radius + 2.0
    for (a1, a2) in sa:
        if (min(math.hypot(a1[0] - pt[0], a1[1] - pt[1]),
                math.hypot(a2[0] - pt[0], a2[1] - pt[1])) > r2):
            continue
        for (b1, b2) in sb:
            if (min(math.hypot(b1[0] - pt[0], b1[1] - pt[1]),
                    math.hypot(b2[0] - pt[0], b2[1] - pt[1])) > r2):
                continue
            if _seg_dist_pair(a1, a2, b1, b2) <= 0.001:
                return True
    return False


def _cluster(gaps, dist_mm=0.3):
    """同一处最窄通道只保留一个代表（距离最近的那个），其余记为重复。"""
    kept = []
    for g in gaps:
        p = g["pt"]
        hit = None
        for k in kept:
            if math.hypot(k["pt"][0] - p[0], k["pt"][1] - p[1]) <= dist_mm:
                hit = k
                break
        if hit is None:
            g["cluster_n"] = 1
            kept.append(g)
        else:
            hit["cluster_n"] = hit.get("cluster_n", 1) + 1
    return kept


def analyze_gaps(geo, thr_mm, eps_mm=EPS_MM, cell=0.5, join_radius=0.8,
                 cluster_mm=0.3):
    """线段-线段近缝：两条线最近距离在 (eps, thr] 之间却没相交。

    这是「智能填充串区」的主因——交点错开零点几丝，填充从针眼漏到隔壁。
    """
    segs = []
    for layer, paths in geo.items():
        for pi, p in enumerate(paths):
            pts = p["pts"]
            key = (layer, tuple(p["addr"]), p["sub"])
            for i in range(len(pts) - 1):
                a, b = pts[i], pts[i + 1]
                if math.hypot(b[0] - a[0], b[1] - a[1]) > 1e-6:
                    segs.append((key, a, b))
    if not segs:
        return []

    by_key = {}
    for (k, a, b) in segs:
        by_key.setdefault(k, []).append((a, b))

    grid = {}
    for idx, (_k, a, b) in enumerate(segs):
        x0, x1 = (a[0], b[0]) if a[0] <= b[0] else (b[0], a[0])
        y0, y1 = (a[1], b[1]) if a[1] <= b[1] else (b[1], a[1])
        for gx in range(int(x0 // cell), int(x1 // cell) + 1):
            for gy in range(int(y0 // cell), int(y1 // cell) + 1):
                grid.setdefault((gx, gy), []).append(idx)

    out = []
    seen = set()
    for ids in grid.values():
        n = len(ids)
        for i in range(n):
            for j in range(i + 1, n):
                ia, ib = ids[i], ids[j]
                if ia > ib:
                    ia, ib = ib, ia
                if (ia, ib) in seen:
                    continue
                seen.add((ia, ib))
                ka, a1, a2 = segs[ia]
                kb, b1, b2 = segs[ib]
                if ka == kb:
                    continue
                d = _seg_dist_pair(a1, a2, b1, b2)
                if d <= eps_mm or d > thr_mm:
                    continue
                pa, pb, _ = closest_pair(a1, a2, b1, b2)
                v1 = (a2[0] - a1[0], a2[1] - a1[1])
                v2 = (b2[0] - b1[0], b2[1] - b1[1])
                l1 = math.hypot(*v1) or 1e-9
                l2 = math.hypot(*v2) or 1e-9
                cos = abs(v1[0] * v2[0] + v1[1] * v2[1]) / (l1 * l2)
                ang = math.degrees(math.acos(max(-1.0, min(1.0, cos))))
                la = math.hypot(a2[0] - a1[0], a2[1] - a1[1])
                lb = math.hypot(b2[0] - b1[0], b2[1] - b1[1])
                # 吸附在哪条线上：挑「这一段更长」的那条（同样位移下角度变化更小）
                if la >= lb:
                    snap_on = "a"
                    snap_layer, snap_addr, snap_sub = ka[0], list(ka[1]), ka[2]
                    snap_from, snap_to = pa, pb
                else:
                    snap_on = "b"
                    snap_layer, snap_addr, snap_sub = kb[0], list(kb[1]), kb[2]
                    snap_from, snap_to = pb, pa
                low = _already_joined(by_key, ka, kb, ((pa[0] + pb[0]) / 2,
                                                       (pa[1] + pb[1]) / 2), join_radius)
                out.append({
                    "kind": "gap",
                    "low_risk": low,
                    "a_layer": ka[0], "a_addr": list(ka[1]), "a_sub": ka[2],
                    "b_layer": kb[0], "b_addr": list(kb[1]), "b_sub": kb[2],
                    "snap_on": snap_on,
                    "snap_layer": snap_layer, "snap_addr": snap_addr, "snap_sub": snap_sub,
                    "snap_from": (round(snap_from[0], 4), round(snap_from[1], 4)),
                    "snap_to": (round(snap_to[0], 4), round(snap_to[1], 4)),
                    "a_pt": (round(pa[0], 4), round(pa[1], 4)),
                    "b_pt": (round(pb[0], 4), round(pb[1], 4)),
                    "dist_mm": round(d, 4),
                    "angle": round(ang, 1),
                    "parallel": ang < 5.0,
                    "layer": ka[0],
                    "pt": (round((pa[0] + pb[0]) / 2, 4), round((pa[1] + pb[1]) / 2, 4)),
                    "end": -1,
                    "mode": "bridge",
                    "extend_to": (round(pb[0], 4), round(pb[1], 4)),
                    "extend_len_mm": round(d, 4),
                    "reason": "",
                })
    out.sort(key=lambda x: x["dist_mm"])
    return _cluster(out, cluster_mm)


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
                # 沿「前一个节点 → 端点」的射线找**最先**碰到的线段；
                # 只接受 ≤ 阈值 的延伸距离 —— 碰到就截止，绝不无限延伸。
                hit = None
                first_t = None
                for (rl, raddr, ridx, a, b) in refs:
                    if rl == lname and raddr == tuple(sp["addr"]) and ridx == idx:
                        continue
                    t = ray_seg_hit(e, (dx, dy), a, b)
                    if t is None:
                        continue
                    if first_t is None or t < first_t:
                        first_t = t
                    if t > thr_mm:
                        continue
                    if hit is None or t < hit[0]:
                        hit = (t, rl, a, b)
                extend_to = None
                extend_len = None
                mode = ""
                reason = ""
                ref_layer = best[1]
                ref_addr = best[2]
                if hit is not None:
                    extend_to = (e[0] + dx * hit[0], e[1] + dy * hit[0])
                    extend_len = hit[0]
                    ref_layer = hit[1]
                    mode = "ray"
                else:
                    # 顺着方向会超出上限（或根本打不到）：改用**垂线段**补齐。
                    # 垂足到端点的距离就是最近距离 d（≤ 上限），必然合规。
                    fx, fy = foot_on_seg(e, best[4], best[5])
                    extend_to = (fx, fy)
                    extend_len = math.hypot(fx - e[0], fy - e[1])
                    mode = "perp"
                    if first_t is not None and first_t > thr_mm:
                        reason = ("方向偏了（要延伸 %.2fmm），已改用垂线补齐"
                                  % first_t)
                    else:
                        reason = "方向打不到线，已改用垂线补齐"
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
                    "reason": reason,
                    "mode": mode,
                })
    issues.sort(key=lambda x: (x["layer"], x["dist_mm"]))
    return issues
