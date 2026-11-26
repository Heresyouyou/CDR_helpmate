# -*- coding: utf-8 -*-
"""平面图分面：把线条当图的边，求“面”，再找每个面的最窄通道。

为什么这样判定：智能填色按「面」工作——一条区域只有完全被轮廓围住才成面。
若两条线之间有一处窄颈（最窄距离 < 阈值），本来分开的两个面就被焊成一个面，
于是填色从窄颈漏到隔壁。所以判据不是“有没有缝”，而是
「这个面在自己身上有没有一处宽度小于阈值的窄颈」。
"""
import math

EPS = 1e-9
NODE_TOL = 3          # 节点坐标取整位数（0.001mm）


def _key(x, y):
    return (round(x, NODE_TOL), round(y, NODE_TOL))


def _seg_intersection(a1, a2, b1, b2):
    """返回 (t, u) 或 None（平行/不相交）。"""
    dx1, dy1 = a2[0] - a1[0], a2[1] - a1[1]
    dx2, dy2 = b2[0] - b1[0], b2[1] - b1[1]
    den = dx1 * dy2 - dy1 * dx2
    if abs(den) < 1e-14:
        return None
    t = ((b1[0] - a1[0]) * dy2 - (b1[1] - a1[1]) * dx2) / den
    u = ((b1[0] - a1[0]) * dy1 - (b1[1] - a1[1]) * dx1) / den
    if -1e-9 <= t <= 1 + 1e-9 and -1e-9 <= u <= 1 + 1e-9:
        return t, u
    return None


def _seg_dist(a1, a2, b1, b2):
    def pt(p, s1, s2):
        vx, vy = s2[0] - s1[0], s2[1] - s1[1]
        l2 = vx * vx + vy * vy
        if l2 <= 1e-18:
            return math.hypot(p[0] - s1[0], p[1] - s1[1])
        t = ((p[0] - s1[0]) * vx + (p[1] - s1[1]) * vy) / l2
        t = 0.0 if t < 0 else (1.0 if t > 1 else t)
        return math.hypot(p[0] - (s1[0] + t * vx), p[1] - (s1[1] + t * vy))
    return min(pt(a1, b1, b2), pt(a2, b1, b2), pt(b1, a1, a2), pt(b2, a1, a2))


def _closest(a1, a2, b1, b2):
    def foot(p, s1, s2):
        vx, vy = s2[0] - s1[0], s2[1] - s1[1]
        l2 = vx * vx + vy * vy
        if l2 <= 1e-18:
            return s1
        t = ((p[0] - s1[0]) * vx + (p[1] - s1[1]) * vy) / l2
        t = 0.0 if t < 0 else (1.0 if t > 1 else t)
        return (s1[0] + t * vx, s1[1] + t * vy)
    cands = [(a1, foot(a1, b1, b2)), (a2, foot(a2, b1, b2)),
             (foot(b1, a1, a2), b1), (foot(b2, a1, a2), b2)]
    return min(cands, key=lambda pq: math.hypot(pq[0][0] - pq[1][0], pq[0][1] - pq[1][1]))


def build_arrangement(geo, cell=1.0):
    """把折线切成以交点分段的边。返回 (nodes, edges)。

    edges: [(u, v, owner_key, ax, ay, bx, by)]
    """
    raw = []          # (a, b, owner)
    for layer, paths in geo.items():
        for p in paths:
            pts = p["pts"]
            owner = (layer, tuple(p["addr"]), p["sub"])
            for i in range(len(pts) - 1):
                a, b = pts[i], pts[i + 1]
                if math.hypot(b[0] - a[0], b[1] - a[1]) > 1e-7:
                    raw.append((a, b, owner))
    n = len(raw)
    splits = [[] for _ in range(n)]

    grid = {}
    for i, (a, b, _o) in enumerate(raw):
        x0, x1 = (a[0], b[0]) if a[0] <= b[0] else (b[0], a[0])
        y0, y1 = (a[1], b[1]) if a[1] <= b[1] else (b[1], a[1])
        for gx in range(int(x0 // cell), int(x1 // cell) + 1):
            for gy in range(int(y0 // cell), int(y1 // cell) + 1):
                grid.setdefault((gx, gy), []).append(i)

    seen = set()
    for ids in grid.values():
        m = len(ids)
        for i in range(m):
            for j in range(i + 1, m):
                ia, ib = ids[i], ids[j]
                if ia > ib:
                    ia, ib = ib, ia
                if (ia, ib) in seen:
                    continue
                seen.add((ia, ib))
                a1, a2, _ = raw[ia]
                b1, b2, _b = raw[ib]
                hit = _seg_intersection(a1, a2, b1, b2)
                if not hit:
                    continue
                t, u = hit
                if 1e-7 < t < 1 - 1e-7:
                    splits[ia].append(t)
                if 1e-7 < u < 1 - 1e-7:
                    splits[ib].append(u)

    node_id = {}
    nodes = []

    def nid(x, y):
        k = _key(x, y)
        i = node_id.get(k)
        if i is None:
            i = len(nodes)
            node_id[k] = i
            nodes.append((x, y))
        return i

    edges = []
    for i, (a, b, owner) in enumerate(raw):
        ts = sorted(set([0.0, 1.0] + splits[i]))
        dx, dy = b[0] - a[0], b[1] - a[1]
        prev = None
        for t in ts:
            p = (a[0] + dx * t, a[1] + dy * t)
            cur = nid(*p)
            if prev is not None and cur != prev:
                edges.append((prev, cur, owner, nodes[prev][0], nodes[prev][1], p[0], p[1]))
            prev = cur
    return nodes, edges


def enumerate_faces(nodes, edges):
    """半边结构遍历所有面，返回 [(面积, [edge_index,...])]。"""
    adj = {}
    for ei, e in enumerate(edges):
        u, v = e[0], e[1]
        adj.setdefault(u, []).append((ei, 0, v))     # 0: u->v
        adj.setdefault(v, []).append((ei, 1, u))     # 1: v->u

    for node, lst in adj.items():
        px, py = nodes[node]
        lst.sort(key=lambda it: math.atan2(
            (nodes[it[2]][1] - py), (nodes[it[2]][0] - px)))

    def nxt(ei, direction):
        u, v = (edges[ei][0], edges[ei][1]) if direction == 0 else (edges[ei][1], edges[ei][0])
        lst = adj[v]
        # 反向半边在 v 的出边里的位置
        idx = None
        for k, (e2, d2, to) in enumerate(lst):
            if e2 == ei and to == u:
                idx = k
                break
        if idx is None:
            return None
        k2 = (idx - 1) % len(lst)          # 顺时针取前一条 → 面在左侧
        e2, d2, _to = lst[k2]
        return (e2, d2)

    visited = set()
    faces = []
    for ei in range(len(edges)):
        for d in (0, 1):
            if (ei, d) in visited:
                continue
            loop = []
            cur = (ei, d)
            guard = 0
            while cur is not None and cur not in visited and guard < 200000:
                visited.add(cur)
                loop.append(cur)
                cur = nxt(*cur)
                guard += 1
            if len(loop) < 3:
                continue
            area = 0.0
            for (e2, d2) in loop:
                u, v = (edges[e2][0], edges[e2][1]) if d2 == 0 else (edges[e2][1], edges[e2][0])
                area += nodes[u][0] * nodes[v][1] - nodes[v][0] * nodes[u][1]
            faces.append((area / 2.0, [e for e, _d in loop]))
    return faces


def _loop_vertices(nodes, edges, loop_edges):
    """按遍历顺序取出这个面的顶点序列。"""
    pts = []
    for ei in loop_edges:
        u = edges[ei][0]
        pts.append(nodes[u])
    return pts


def _sub_area(pts, i, j):
    """把环从 i 走到 j，加上弦，围出的多边形面积。"""
    n = len(pts)
    poly = []
    k = i
    while True:
        poly.append(pts[k])
        if k == j:
            break
        k = (k + 1) % n
    a = 0.0
    m = len(poly)
    for t in range(m):
        x1, y1 = poly[t]
        x2, y2 = poly[(t + 1) % m]
        a += x1 * y2 - x2 * y1
    return abs(a) / 2.0


def face_narrowest(nodes, edges, loop_edges, thr_mm, cell=None, eps_mm=0.0):
    """面内所有「宽度在 (eps, thr] 之间」的窄颈。返回 [(d, e1, e2, p, q), ...]。"""
    cell = cell or max(thr_mm * 4, 0.5)
    items = []
    for k, ei in enumerate(loop_edges):
        _u, _v, _o, ax, ay, bx, by = edges[ei]
        items.append((k, ei, ax, ay, bx, by,
                      (ax + bx) / 2.0, (ay + by) / 2.0))
    grid = {}
    for it in items:
        grid.setdefault((int(it[5] // cell), int(it[6] // cell)), []).append(it)
    best = []
    n = len(items)
    for it in items:
        gx, gy = int(it[5] // cell), int(it[6] // cell)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for jt in grid.get((gx + dx, gy + dy), ()):
                    if jt[0] <= it[0]:
                        continue
                    # 只跳过「共享端点」的相邻边；隔着缺口的两段虽然相邻于环上，也不许跳
                    a1, a2 = edges[it[1]][0], edges[it[1]][1]
                    b1, b2 = edges[jt[1]][0], edges[jt[1]][1]
                    if a1 in (b1, b2) or a2 in (b1, b2):
                        continue
                    d = _seg_dist((it[2], it[3]), (it[4], it[5]),
                                  (jt[2], jt[3]), (jt[4], jt[5]))
                    if d <= eps_mm or d > thr_mm:
                        continue
                    p, q = _closest((it[2], it[3]), (it[4], it[5]),
                                    (jt[2], jt[3]), (jt[4], jt[5]))
                    best.append((d, it[1], jt[1], p, q))
    return best


def _make_issue(edges, d, e1, e2, p, q, area):
    o1 = edges[e1][2]
    o2 = edges[e2][2]
    return {
        "kind": "gap",
        "low_risk": False,
        "parallel": False,
        "a_layer": o1[0], "a_addr": list(o1[1]), "a_sub": o1[2],
        "b_layer": o2[0], "b_addr": list(o2[1]), "b_sub": o2[2],
        "a_pt": (round(p[0], 4), round(p[1], 4)),
        "b_pt": (round(q[0], 4), round(q[1], 4)),
        "dist_mm": round(d, 4),
        "angle": 0.0,
        "layer": o1[0],
        "pt": (round((p[0] + q[0]) / 2, 4), round((p[1] + q[1]) / 2, 4)),
        "end": -1, "mode": "bridge",
        "extend_to": (round(q[0], 4), round(q[1], 4)),
        "extend_len_mm": round(d, 4),
        "reason": "面内窄颈 %.4fmm" % d,
        "face_area": round(area, 1),
        "cluster_n": 1,
        "snap_on": "a", "snap_layer": o1[0], "snap_addr": list(o1[1]),
        "snap_sub": o1[2], "snap_from": (round(p[0], 4), round(p[1], 4)),
        "snap_to": (round(q[0], 4), round(q[1], 4)),
    }


def find_pinches(geo, thr_mm, eps_mm=0.005, min_area=0.5, min_region=1.0):
    """对每个面找狭窄通道：一个面里可能有多处，各自独立报。

    判据来自智能填色的原理：面被轮廓围住才成面；若一个面内部存在宽度
    小于阈值的窄颈，说明那里本该是两堵墙，现在焊成了一体 → 需要分割。
    """
    nodes, edges = build_arrangement(geo)
    faces = enumerate_faces(nodes, edges)
    out = []
    for area, loop in faces:
        if area <= min_area or area > 1e7:
            continue
        cands = face_narrowest(nodes, edges, loop, thr_mm,
                               cell=max(thr_mm * 4, 0.5), eps_mm=eps_mm)
        if not cands:
            continue
        cands.sort(key=lambda t: t[0])
        verts = _loop_vertices(nodes, edges, loop)
        pos = {ei: k for k, ei in enumerate(loop)}
        picked = []
        for d, e1, e2, p, q in cands:
            if any(math.hypot(p[0] - pp[0], p[1] - pp[1]) <= 0.3 for pp in picked):
                continue
            if edges[e1][2] == edges[e2][2]:
                continue
            # 只有「在这里切一刀会分出两块像样的区域」才算真分隔；
            # 切开后一边只是细条/尖角的，修了也没意义。
            try:
                a1 = _sub_area(verts, pos[e1], pos[e2])
                a2 = abs(area) - a1
                if min(a1, a2) < min_region:
                    continue
            except Exception:
                pass
            picked.append(p)
            out.append(_make_issue(edges, d, e1, e2, p, q, area))
    out.sort(key=lambda x: x["dist_mm"])
    return out, len(faces)
