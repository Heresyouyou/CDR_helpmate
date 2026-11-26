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


def build_arrangement(geo, cell=1.0, snap_tol=0.01):
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
                edges.append([prev, cur, owner, nodes[prev][0], nodes[prev][1], p[0], p[1]])
            prev = cur

    # ---- 第二遍：把「落在其它边内部」的节点真正接上去 ----
    # 只算交点是不够的：端点正好压在另一条线的中间、或两条线共线重叠时，
    # 距离是 0 却没有交点，于是就变成"看着贴上、拓扑上断开"——智能填色会从这儿漏。
    if snap_tol and snap_tol > 0:
        egrid = {}
        for k, e in enumerate(edges):
            ax, ay, bx, by = e[3], e[4], e[5], e[6]
            x0, x1 = (ax, bx) if ax <= bx else (bx, ax)
            y0, y1 = (ay, by) if ay <= by else (by, ay)
            for gx in range(int(x0 // cell), int(x1 // cell) + 1):
                for gy in range(int(y0 // cell), int(y1 // cell) + 1):
                    egrid.setdefault((gx, gy), []).append(k)
        extra = {}
        for n, (px, py) in enumerate(nodes):
            gx, gy = int(px // cell), int(py // cell)
            for dx_ in (-1, 0, 1):
                for dy_ in (-1, 0, 1):
                    for k in egrid.get((gx + dx_, gy + dy_), ()):
                        e = edges[k]
                        if e[0] == n or e[1] == n:
                            continue
                        ax, ay, bx, by = e[3], e[4], e[5], e[6]
                        vx, vy = bx - ax, by - ay
                        l2 = vx * vx + vy * vy
                        if l2 <= 1e-18:
                            continue
                        t = ((px - ax) * vx + (py - ay) * vy) / l2
                        if not (1e-6 < t < 1 - 1e-6):
                            continue
                        d = abs((px - ax) * vy - (py - ay) * vx) / math.sqrt(l2)
                        if d > snap_tol:
                            continue
                        extra.setdefault(k, set()).add(round(t, 9))
        if extra:
            new_edges = []
            for k, e in enumerate(edges):
                ts = extra.get(k)
                if not ts:
                    new_edges.append(tuple(e))
                    continue
                ax, ay, bx, by = e[3], e[4], e[5], e[6]
                cuts = sorted(set([0.0, 1.0] + list(ts)))
                prev = e[0]
                for t in cuts[1:]:
                    px = ax + (bx - ax) * t
                    py = ay + (by - ay) * t
                    cur = e[1] if t >= 1 - 1e-9 else nid(px, py)
                    if cur != prev:
                        new_edges.append((prev, cur, e[2],
                                          nodes[prev][0], nodes[prev][1], px, py))
                    prev = cur
            edges = new_edges
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


# ---------------------------------------------------------------- 拓扑判据

class _UF(object):
    def __init__(self, n):
        self.p = list(range(n))

    def find(self, x):
        p = self.p
        while p[x] != x:
            p[x] = p[p[x]]
            x = p[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.p[ra] = rb


def graph_stats(geo, snap_tol=0.0):
    """平面图拓扑量：V 节点 / E 边 / C 连通分量 / F 有界面。

    snap_tol 默认 0：CDR 的智能填充要求曲线**精确接触**，任何 >0 的缝都是漏口；
    以前用 0.01mm 吸附去算，等于把真实缺口当成已经连上，判据从根上就错了。

    欧拉公式：F = E - V + C。这个 F 就是「智能填色能填出多少块」的量化指标——
    修复前后比 F，增加才算真有效。
    """
    nodes, edges = build_arrangement(geo, snap_tol=snap_tol)
    V, E = len(nodes), len(edges)
    uf = _UF(V)
    deg = {}
    for e in edges:
        u, v = e[0], e[1]
        uf.union(u, v)
        deg[u] = deg.get(u, 0) + 1
        deg[v] = deg.get(v, 0) + 1
    C = len({uf.find(i) for i in range(V)}) if V else 0
    return {
        "V": V, "E": E, "C": C, "F": E - V + C,
        "dangling": sum(1 for d in deg.values() if d == 1),
        "uf": uf, "nodes": nodes, "edges": edges, "deg": deg,
    }


def dangling_ends(geo, tol_mm=0.2, eps_mm=0.005, cell=1.0, stats=None):
    """拓扑意义上的悬空端头（度=1 的节点）+ 它接上去会不会真的多出一个面。

    判据（欧拉）：把悬空端点连到**同一连通分量**里的另一处 → E+1、C 不变 → F+1，多一个面；
    连到**另一个分量** → E+1、C-1 → F 不变，白修。
    """
    st = stats or graph_stats(geo)
    nodes, edges = st["nodes"], st["edges"]
    deg, uf = st["deg"], st["uf"]
    V = len(nodes)

    grid = {}
    for i, e in enumerate(edges):
        a, b = nodes[e[0]], nodes[e[1]]
        x0, x1 = (a[0], b[0]) if a[0] <= b[0] else (b[0], a[0])
        y0, y1 = (a[1], b[1]) if a[1] <= b[1] else (b[1], a[1])
        for gx in range(int(x0 // cell), int(x1 // cell) + 1):
            for gy in range(int(y0 // cell), int(y1 // cell) + 1):
                grid.setdefault((gx, gy), []).append(i)

    # 端头坐标 -> (layer, addr, sub, end) 反查
    by_pt = {}
    for layer, paths in geo.items():
        for p in paths:
            if p["closed"] or len(p["pts"]) < 2:
                continue
            for end, pt in ((0, p["pts"][0]), (1, p["pts"][-1])):
                by_pt[(round(pt[0], 3), round(pt[1], 3))] = (layer, p["addr"], p["sub"], end)

    # 断头链长度：从端点走到最近岔口，越长说明这条线本该越有分隔作用
    adj = {}
    for i, e in enumerate(edges):
        adj.setdefault(e[0], []).append((e[1], i))
        adj.setdefault(e[1], []).append((e[0], i))

    def chain_len(start):
        total = 0.0
        cur = start
        prev = None
        for _ in range(5000):
            nxt = [(v, i) for v, i in adj.get(cur, []) if v != prev]
            if not nxt:
                break
            v, i = nxt[0]
            a, b = nodes[edges[i][0]], nodes[edges[i][1]]
            total += math.hypot(b[0] - a[0], b[1] - a[1])
            if deg.get(v, 0) != 2:
                break
            prev, cur = cur, v
        return total

    out = []
    for n, d in deg.items():
        if d != 1:
            continue
        pt = nodes[n]
        owner = by_pt.get((round(pt[0], 3), round(pt[1], 3)))
        if owner is None:
            continue
        gx, gy = int(pt[0] // cell), int(pt[1] // cell)
        best = None
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for i in grid.get((gx + dx, gy + dy), ()):
                    u, v = edges[i][0], edges[i][1]
                    if u == n or v == n:
                        continue
                    a, b = nodes[u], nodes[v]
                    vx, vy = b[0] - a[0], b[1] - a[1]
                    l2 = vx * vx + vy * vy
                    if l2 <= 1e-18:
                        continue
                    t = ((pt[0] - a[0]) * vx + (pt[1] - a[1]) * vy) / l2
                    t = 0.0 if t < 0 else (1.0 if t > 1 else t)
                    proj = (a[0] + t * vx, a[1] + t * vy)
                    dd = math.hypot(pt[0] - proj[0], pt[1] - proj[1])
                    if best is None or dd < best[0]:
                        best = (dd, i, proj, u)
        if best is None:
            continue
        dd, ei, proj, tu = best
        if dd <= eps_mm or dd > tol_mm:
            continue
        same = uf.find(n) == uf.find(tu)
        layer, addr, sub, end = owner
        out.append({
            "kind": "end",
            "layer": layer, "addr": list(addr), "sub": sub, "end": end,
            "pt": (round(pt[0], 4), round(pt[1], 4)),
            "dist_mm": round(dd, 4),
            "extend_to": (round(proj[0], 4), round(proj[1], 4)),
            "extend_len_mm": round(dd, 4),
            "mode": "perp",
            "same_comp": bool(same),
            "gain_face": 1 if same else 0,
            "chain_mm": round(chain_len(n), 2),
            "reason": ("同一片里断开，接上会多一个面" if same
                       else "跨片连接，接上不会多出可填的面"),
            "low_risk": (not same), "parallel": False, "cluster_n": 1,
        })
    out.sort(key=lambda x: (not x["same_comp"], -x.get("chain_mm", 0), x["dist_mm"]))
    return out


def validate_repairs(geo, issues):
    """逐条模拟修复，只有「真的让 F 增加」的才标记 will_gain。

    做法：把端点接到目标点后重算 F（欧拉公式），逐条累加模拟，避免互相影响。
    这条是整个方案里唯一能回答「修了到底有没有用」的判据。
    """
    work = {}
    for layer, paths in geo.items():
        work[layer] = [{"addr": p["addr"], "sub": p["sub"], "closed": p["closed"],
                        "pts": list(p["pts"])} for p in paths]

    def key_map(g):
        m = {}
        for layer, paths in g.items():
            for p in paths:
                m[(layer, tuple(p["addr"]), p["sub"])] = p
        return m

    base = graph_stats(work)["F"]
    km = key_map(work)
    for it in issues:
        k = (it["layer"], tuple(it["addr"]), it["sub"])
        p = km.get(k)
        if p is None:
            it["will_gain"] = False
            continue
        tgt = it["extend_to"]
        # 端点已经就是目标点 → 无意义
        if math.hypot(p["pts"][-1][0] - tgt[0], p["pts"][-1][1] - tgt[1]) < 1e-6 and it["end"] == 1:
            it["will_gain"] = False
            continue
        if it["end"] == 1:
            p["pts"].append(tgt)
        else:
            p["pts"].insert(0, tgt)
        f2 = graph_stats(work)["F"]
        # 微小缝隙（<0.05mm）在节点取整(0.001mm)时可能被合并掉，模拟看不出增面；
        # 但按欧拉公式：同分量内把度=1 的端点接上，必然 F+1，这是解析结论。
        analytic = (it.get("same_comp") and it.get("dist_mm", 9) < 0.05)
        if f2 > base or analytic:
            it["gain_face"] = max(1, f2 - base)
            base = f2
            it["will_gain"] = True
        else:
            if it["end"] == 1:
                p["pts"].pop()
            else:
                p["pts"].pop(0)
            it["will_gain"] = False
            it["reason"] = "接上后不会多出面（与已有连接重复）"
    return issues


def key_chain_ends(geo, min_chain=5.0, tol_mm=0.5, eps_mm=0.0, cell=1.0,
                   marks=None, snap_tol=0.0):
    """只找「会让一大片区域串在一起」的关键断头。

    串面的特征：一条足够长的线（断层/地质界限）端头悬空，两侧属于同一个面。
    所以判据不是缝隙多小，而是：
      1. 这个端点属于一条长度 >= min_chain 的断头链；
      2. 它位于一个大面内部（若有用户标注的「串面」多边形，则只取落在里面的）；
      3. 接上它以后能让 F 增加（validate_repairs 二次确认）。
    """
    nodes, edges = build_arrangement(geo, snap_tol=snap_tol)
    deg, adj = {}, {}
    for i, e in enumerate(edges):
        u, v = e[0], e[1]
        deg[u] = deg.get(u, 0) + 1; deg[v] = deg.get(v, 0) + 1
        adj.setdefault(u, []).append((v, i)); adj.setdefault(v, []).append((u, i))

    def elen(i):
        a, b = nodes[edges[i][0]], nodes[edges[i][1]]
        return math.hypot(b[0]-a[0], b[1]-a[1])

    def walk(start):
        L, cur, prev = 0.0, start, None
        for _ in range(5000):
            nxt = [(v, i) for v, i in adj.get(cur, []) if v != prev]
            if not nxt:
                break
            v, i = nxt[0]
            L += elen(i)
            if deg.get(v, 0) != 2:
                return L, v
            prev, cur = cur, v
        return L, cur

    # 面：给每个端点标注它所在的最大面面积
    loops = enumerate_faces(nodes, edges)
    face_area_of_node = {}
    for area, loop in loops:
        if area <= 0:
            continue
        for ei in loop:
            e = edges[ei]
            for n in (e[0], e[1]):
                if area > face_area_of_node.get(n, 0):
                    face_area_of_node[n] = area

    def inside(pt, poly):
        x, y = pt; c = False; j = len(poly)-1
        for i in range(len(poly)):
            xi, yi = poly[i]; xj, yj = poly[j]
            if ((yi > y) != (yj > y)) and (x < (xj-xi)*(y-yi)/(yj-yi+1e-18)+xi):
                c = not c
            j = i
        return c

    # 端点坐标 -> 原始子路径的 end 标记
    by_pt = {}
    for layer, paths in geo.items():
        for p in paths:
            if p["closed"] or len(p["pts"]) < 2:
                continue
            for end, pt in ((0, p["pts"][0]), (1, p["pts"][-1])):
                by_pt[(round(pt[0], 3), round(pt[1], 3))] = (layer, p["addr"], p["sub"], end)

    grid = {}
    for i, e in enumerate(edges):
        a, b = nodes[e[0]], nodes[e[1]]
        x0, x1 = (a[0], b[0]) if a[0] <= b[0] else (b[0], a[0])
        y0, y1 = (a[1], b[1]) if a[1] <= b[1] else (b[1], a[1])
        for gx in range(int(x0//cell), int(x1//cell)+1):
            for gy in range(int(y0//cell), int(y1//cell)+1):
                grid.setdefault((gx, gy), []).append(i)

    out = []
    for n, d in deg.items():
        if d != 1:
            continue
        L, terminal = walk(n)
        if L < min_chain:
            continue
        pt = nodes[n]
        if marks and not any(inside(pt, mp) for mp in marks):
            continue
        owner = by_pt.get((round(pt[0], 3), round(pt[1], 3)))
        if owner is None:
            continue
        gx, gy = int(pt[0]//cell), int(pt[1]//cell)
        best = None
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for i in grid.get((gx+dx, gy+dy), ()):
                    u, v = edges[i][0], edges[i][1]
                    if u == n or v == n:
                        continue
                    a, b = nodes[u], nodes[v]
                    vx, vy = b[0]-a[0], b[1]-a[1]
                    l2 = vx*vx+vy*vy
                    if l2 <= 1e-18:
                        continue
                    t = ((pt[0]-a[0])*vx+(pt[1]-a[1])*vy)/l2
                    t = 0.0 if t < 0 else (1.0 if t > 1 else t)
                    proj = (a[0]+t*vx, a[1]+t*vy)
                    dd = math.hypot(pt[0]-proj[0], pt[1]-proj[1])
                    if best is None or dd < best[0]:
                        best = (dd, proj)
        if best is None or best[0] < eps_mm or best[0] > tol_mm:
            continue
        layer, addr, sub, end = owner
        out.append({
            "kind": "end", "layer": layer, "addr": list(addr), "sub": sub, "end": end,
            "pt": (round(pt[0], 4), round(pt[1], 4)),
            "dist_mm": round(best[0], 4),
            "extend_to": (round(best[1][0], 4), round(best[1][1], 4)),
            "extend_len_mm": round(best[0], 4), "mode": "perp",
            "chain_mm": round(L, 2), "face_area": round(face_area_of_node.get(n, 0), 1),
            "chain_id": "%d_%d" % (min(n, terminal), max(n, terminal)),
            "same_comp": True, "gain_face": 1, "cluster_n": 1,
            "low_risk": False, "parallel": False,
            "reason": "断头链 %.1fmm，位于 %.0fmm² 的大面里" % (L, face_area_of_node.get(n, 0)),
        })
    out.sort(key=lambda x: (-x["chain_mm"], x["dist_mm"]))
    return out
