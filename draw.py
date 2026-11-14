#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""结单工单分层配额抽样。

用法示例：
    python3 draw.py --T 2026-11-10T18:00:00 \
        --quota app=40,chat=25,store=12,web=14 \
        --xlsx closed.xlsx --weights weights.txt \
        --out-sample sample.csv --out-report report.txt

规则要点：
  * 抽样框：closed_at <= T 的工单。T 是“截止时刻”，截止瞬间结单的工单
    （closed_at == T）属于框内。
  * 各渠道候选按 (closed_at, tid) 升序排列。
  * 每层按配额取排序序列的前 quota 个位置作为“抽中”；其中 refunded=是
    的工单由同层此后的候选依次顶替（候补若也是退款者则继续向后找）。
  * 同层耗尽后，按 quota 参数中的渠道次序向其余层征借；只有自身配额
    已补足、且候选序列仍有可用工单的层可出借。借入工单的 channel 写
    来源渠道，layer 写目标层。
  * 替换深度 = 为取得顶替工单而依次尝试的候选个数（含被跳过的退款者）。
"""

import argparse
import csv
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, timedelta

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REF_RE = re.compile(r"([A-Z]+)(\d+)$")


# ---------------------------------------------------------------- xlsx reader
def _col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


def _cell_text(cell, shared):
    """Return the displayed text of one <c> element."""
    t = cell.get("t")
    if t == "inlineStr":
        return "".join(x.text or "" for x in cell.iter(NS + "t"))
    v = cell.find(NS + "v")
    if v is None or v.text is None:
        return ""
    if t == "s":                      # shared string index
        try:
            return shared[int(v.text)]
        except (ValueError, IndexError):
            return ""
    if t == "b":
        return "TRUE" if v.text in ("1", "true") else "FALSE"
    return v.text


def read_xlsx_rows(path):
    """读取第一个 worksheet，返回 {row_no: {col_idx: text}}。"""
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sroot.iter(NS + "si"):
                shared.append("".join(x.text or "" for x in si.iter(NS + "t")))

        sheets = sorted(
            (n for n in z.namelist()
             if re.match(r"xl/worksheets/sheet\d+\.xml$", n)),
            key=lambda n: int(re.search(r"sheet(\d+)", n).group(1)),
        )
        if not sheets:
            raise SystemExit("错误：xlsx 中没有 worksheet")
        root = ET.fromstring(z.read(sheets[0]))

    rows = {}
    for row in root.iter(NS + "row"):
        rno = int(row.get("r"))
        cells = {}
        for cell in row.iter(NS + "c"):
            ref = cell.get("r")
            m = REF_RE.match(ref or "")
            if not m:
                continue
            cells[_col_index(m.group(1))] = _cell_text(cell, shared)
        rows[rno] = cells
    return rows


# ------------------------------------------------------------------- parsing
def norm_ts(s):
    """Parse a timestamp; return datetime or None."""
    s = (s or "").strip().replace("/", "-")
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace(" ", "T"))
    except ValueError:
        pass
    try:
        return datetime.strptime(s, "%Y-%m-%d")
    except ValueError:
        pass
    # Excel serial date (days since 1899-12-30), as a fallback
    try:
        f = float(s)
        if 1 <= f <= 100000:
            return datetime(1899, 12, 30) + timedelta(days=f)
    except ValueError:
        pass
    return None


def parse_sat(s):
    """Return a normalized sat value (int/float) or None if unanswered."""
    s = (s or "").strip()
    if not s:
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    return int(f) if f == int(f) else f


REFUND_TRUE = {"是", "true", "TRUE", "True", "1", "yes", "Y", "y"}


def is_refunded(s):
    return (s or "").strip() in REFUND_TRUE


def parse_quota(s):
    """Parse --quota, preserving the parameter's channel order."""
    q = {}
    for part in re.split(r"[,;]", s or ""):
        part = part.strip()
        if not part:
            continue
        if "=" not in part:
            raise SystemExit(f"错误：--quota 片段无法解析：{part!r}")
        k, v = part.split("=", 1)
        k = k.strip()
        try:
            n = int(v.strip())
        except ValueError:
            raise SystemExit(f"错误：--quota 配额不是整数：{part!r}")
        if n < 0:
            raise SystemExit(f"错误：--quota 配额不能为负：{part!r}")
        q[k] = n
    if not q:
        raise SystemExit("错误：--quota 为空")
    return q


def parse_weights(path):
    w = {}
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.replace(",", " ").split()
            if len(parts) >= 2:
                try:
                    w[parts[0]] = float(parts[1])
                except ValueError:
                    continue
    return w


# ------------------------------------------------------------ draw primitives
def find_replacement(order, start, taken):
    """从 order 的 start 位置起向后找第一个可用候选。

    返回 (ticket|None, depth, new_pos)。depth 为依次尝试过的候选个数，
    含被跳过的退款者；已占用的工单不再尝试、也不计数。
    """
    i = start
    scanned = 0
    n = len(order)
    while i < n:
        c = order[i]
        i += 1
        if c["tid"] in taken:
            continue
        scanned += 1
        if not c["refunded"]:
            return c, scanned, i
    return None, scanned, i


def remaining_usable(order, pos, taken):
    return sum(1 for c in order[pos:]
               if not c["refunded"] and c["tid"] not in taken)


def fmt_num(x):
    return f"{x:g}"


# ---------------------------------------------------------------------- main
def main(argv=None):
    ap = argparse.ArgumentParser(description="结单工单分层配额抽样")
    ap.add_argument("--T", required=True,
                    help="抽样框截止时刻（含），如 2026-11-10T18:00:00")
    ap.add_argument("--quota", required=True,
                    help="各渠道配额，如 app=40,chat=25,store=12,web=14")
    ap.add_argument("--xlsx", default="closed.xlsx")
    ap.add_argument("--weights", default="weights.txt")
    ap.add_argument("--out-sample", default="sample.csv")
    ap.add_argument("--out-report", default="report.txt")
    args = ap.parse_args(argv)

    quota = parse_quota(args.quota)
    channels = list(quota.keys())          # 渠道次序 = quota 参数次序
    weights = parse_weights(args.weights)
    T = norm_ts(args.T)
    if T is None:
        raise SystemExit(f"错误：无法解析 --T {args.T!r}")

    raw = read_xlsx_rows(args.xlsx)
    rnos = sorted(raw)
    if not rnos:
        raise SystemExit("错误：xlsx 为空")

    hdr_cells = raw[rnos[0]]
    hdr = {}
    for col, text in hdr_cells.items():
        name = (text or "").strip().lower()
        if name:
            hdr[name] = col
    for need in ("tid", "channel", "closed_at", "refunded", "sat"):
        if need not in hdr:
            raise SystemExit(f"错误：表头缺少列 {need!r}")

    frame = []                # 框内工单
    summary_rows = []         # 汇总行（无 tid 或 channel）
    bad_ts_rows = []          # closed_at 无法解析的行
    dup_rows = []             # 重复 tid 中被忽略的行
    seen_tids = set()

    for rno in rnos[1:]:
        cells = raw[rno]
        tid = (cells.get(hdr["tid"], "") or "").strip()
        ch = (cells.get(hdr["channel"], "") or "").strip()
        if not tid or not ch:
            summary_rows.append(rno)
            continue
        ts_raw = (cells.get(hdr["closed_at"], "") or "").strip()
        ts = norm_ts(ts_raw)
        if ts is None:
            bad_ts_rows.append(rno)
            continue
        if tid in seen_tids:
            dup_rows.append(rno)
            continue
        seen_tids.add(tid)
        # 截止时刻是闭区间：恰在 T 结单的工单入框。
        if ts <= T:
            frame.append({
                "tid": tid,
                "channel": ch,
                "closed_at": ts_raw,
                "dt": ts,
                "refunded": is_refunded(cells.get(hdr["refunded"], "")),
                "sat": parse_sat(cells.get(hdr["sat"], "")),
            })

    frame_counts = Counter(r["channel"] for r in frame)
    usable_counts = Counter(r["channel"] for r in frame if not r["refunded"])

    # 各层候选：按 (结单时刻, 工单号) 升序。
    orders = {
        ch: sorted((r for r in frame if r["channel"] == ch),
                   key=lambda r: (r["dt"], r["tid"]))
        for ch in channels
    }

    pos = {ch: 0 for ch in channels}
    taken = set()
    slots = []        # 样本条目
    pairs = []        # 替换对（含同层与借层）
    pending = {ch: [] for ch in channels}        # 同层未找到顶替的退款抽中者
    window_short = {ch: 0 for ch in channels}    # 本层候选不足配额之数

    # ---- 第一轮：各层按各自配额抽取，并在同层寻找顶替者 -----------------
    for ch in channels:
        q = quota[ch]
        order = orders[ch]
        win = order[pos[ch]:pos[ch] + q]
        pos[ch] += len(win)
        window_short[ch] = q - len(win)

        # 排序序列的前 q 个位置即抽中；其中可用者直接入选。
        for t in win:
            if t["refunded"]:
                continue
            taken.add(t["tid"])
            slots.append({
                "ticket": t, "layer": ch, "channel": ch,
                "role": "入选", "replaced": None, "depth": 0,
            })

        # 抽中但为退款者：从本层截断位之后依次找候补。
        for t in win:
            if not t["refunded"]:
                continue
            repl, depth, newpos = find_replacement(order, pos[ch], taken)
            if repl is None:
                pos[ch] = newpos
                pending[ch].append(t)
                continue
            pos[ch] = newpos
            taken.add(repl["tid"])
            slots.append({
                "ticket": repl, "layer": ch, "channel": ch,
                "role": "顶替", "replaced": t["tid"], "depth": depth,
            })
            pairs.append({
                "refunded": t["tid"], "replacement": repl["tid"],
                "origin_layer": ch, "source_layer": ch,
                "type": "同层", "depth": depth,
            })

    # ---- 第二轮：同层耗尽后，按渠道次序向其余各层征借 -------------------
    borrow_ops = []
    unresolved = []
    # 各层自给实抽数：只有自身配额已由本层候选补足、且序列仍有可用
    # 工单的层才可出借（靠借入才凑够配额的层不算“已补足自身配额”）。
    own_counts = Counter(s["layer"] for s in slots
                         if s["ticket"]["channel"] == s["layer"])

    for ch in channels:
        need_list = list(pending[ch]) + [None] * window_short[ch]
        for orig in need_list:
            lenders = []                 # 可出借的层（渠道次序）
            for other in channels:
                if other == ch:
                    continue
                # 必须已补足自身配额，且排序序列仍有可用工单。
                if own_counts[other] != quota[other]:
                    continue
                rem = remaining_usable(orders[other], pos[other], taken)
                if rem > 0:
                    lenders.append(other)

            filled = False
            for other in lenders:        # 渠道次序
                before = pos[other]
                repl, depth, newpos = find_replacement(
                    orders[other], before, taken)
                if repl is None:            # 理论上不会发生
                    pos[other] = newpos
                    continue
                pos[other] = newpos
                taken.add(repl["tid"])
                role = "借入顶替" if orig is not None else "借入"
                slots.append({
                    "ticket": repl, "layer": ch, "channel": other,
                    "role": role,
                    "replaced": orig["tid"] if orig else None,
                    "depth": depth,
                })
                if orig is not None:
                    pairs.append({
                        "refunded": orig["tid"], "replacement": repl["tid"],
                        "origin_layer": ch, "source_layer": other,
                        "type": "借层", "depth": depth,
                    })
                borrow_ops.append({
                    "target": ch, "source": other,
                    "replaced_tid": orig["tid"] if orig else "",
                    "replacement_tid": repl["tid"],
                    "depth": depth,
                    "pos_before": before, "pos_after": pos[other],
                })
                filled = True
                break

            if not filled:
                unresolved.append((ch, orig["tid"] if orig else ""))

    # ---- 计数与校验 -----------------------------------------------------
    frame_total = len(frame)
    sample_total = len(slots)
    real_counts = Counter(s["layer"] for s in slots)
    sum_frame_layers = sum(frame_counts.get(ch, 0)
                           for ch in set(frame_counts) | set(channels))
    sum_real_layers = sum(real_counts.get(ch, 0) for ch in channels)
    unique_tids = len({s["ticket"]["tid"] for s in slots})
    bad_refunded = sum(1 for s in slots if s["ticket"]["refunded"])
    bad_time = sum(1 for s in slots if s["ticket"]["dt"] > T)
    quota_sum = sum(quota.values())

    assert frame_total == sum_frame_layers, "框数与各层框数之和不等"
    assert sample_total == sum_real_layers, "样本数与各层实抽之和不等"
    assert unique_tids == sample_total, "样本存在重复 tid"
    assert bad_refunded == 0, "样本中出现退款者"
    assert bad_time == 0, "样本中出现超出截止时刻的工单"
    assert quota_sum == sample_total + len(unresolved), "配额、实抽与缺口不一致"
    depth_dist = Counter(p["depth"] for p in pairs)
    same_pairs = [p for p in pairs if p["type"] == "同层"]
    cross_pairs = [p for p in pairs if p["type"] == "借层"]

    # ---- 加权回访完成率 -------------------------------------------------
    # 权重取工单自身渠道（借入工单的 channel 即来源渠道）；sat 非空记完成。
    w_sum = 0.0
    w_done = 0.0
    for s in slots:
        ch = s["ticket"]["channel"]
        if ch not in weights:
            raise SystemExit(f"错误：weights.txt 缺少渠道 {ch!r} 的权重")
        w = weights[ch]
        w_sum += w
        if s["ticket"]["sat"] is not None:
            w_done += w
    rate = (100.0 * w_done / w_sum) if w_sum else 0.0

    borrow_in = Counter(b["target"] for b in borrow_ops)
    borrow_out = Counter(b["source"] for b in borrow_ops)

    # ---- sample.csv -----------------------------------------------------
    fields = ["tid", "channel", "closed_at", "refunded", "sat", "layer"]
    # 目标层按渠道次序；层内按结单时刻、工单号。
    layer_rank = {ch: i for i, ch in enumerate(channels)}
    out_slots = sorted(
        slots,
        key=lambda s: (layer_rank[s["layer"]],
                       s["ticket"]["dt"], s["ticket"]["tid"]),
    )
    with open(args.out_sample, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for s in out_slots:
            t = s["ticket"]
            w.writerow({
                "tid": t["tid"],
                "channel": t["channel"],
                "closed_at": t["closed_at"],
                "refunded": "是" if t["refunded"] else "否",
                "sat": "" if t["sat"] is None else fmt_num(t["sat"]),
                "layer": s["layer"],
            })

    # 回读 sample.csv，确认“框 / 样本 / 报告”三方计数一致。
    with open(args.out_sample, encoding="utf-8") as f:
        csv_rows = sum(1 for _ in csv.DictReader(f))
    assert csv_rows == sample_total == sum_real_layers, \
        "框、样本、报告三方计数不一致"

    # ---- report.txt -----------------------------------------------------
    with open(args.out_report, "w", encoding="utf-8") as f:
        def a(line):
            f.write(line + "\n")

        a(f"frame total={frame_total} data_rows={len(raw) - 1 - len(summary_rows)} "
          f"summary_rows={len(summary_rows)} bad_ts_rows={len(bad_ts_rows)} "
          f"dup_tid_rows={len(dup_rows)} cutoff={args.T} "
          f"cutoff_inclusive=yes")
        for ch in channels:
            a(f"frame layer={ch} count={frame_counts.get(ch, 0)} "
              f"usable={usable_counts.get(ch, 0)} "
              f"refunded={frame_counts.get(ch, 0) - usable_counts.get(ch, 0)}")
        for ch in channels:
            a(f"quota layer={ch} quota={quota[ch]}")
        for ch in channels:
            a(f"drawn layer={ch} drawn={real_counts.get(ch, 0)} "
              f"borrowed_in={borrow_in.get(ch, 0)} "
              f"borrowed_out={borrow_out.get(ch, 0)} "
              f"gap={quota[ch] - real_counts.get(ch, 0)}")
        a(f"pairs total={len(pairs)} same={len(same_pairs)} "
          f"cross={len(cross_pairs)}")
        for p in pairs:
            a(f"pair replaced={p['refunded']} replacement={p['replacement']} "
              f"depth={p['depth']} kind={p['type']} "
              f"target_layer={p['origin_layer']} "
              f"source_channel={p['source_layer']}")
        if depth_dist:
            for d in sorted(depth_dist):
                a(f"depth value={d} count={depth_dist[d]}")
        else:
            a("depth value=0 count=0")
        if borrow_ops:
            for b in borrow_ops:
                a(f"borrow target_layer={b['target']} "
                  f"source_layer={b['source']} "
                  f"replaced={b['replaced_tid'] or '-'} "
                  f"replacement={b['replacement_tid']} "
                  f"depth={b['depth']} "
                  f"source_pos_before={b['pos_before']} "
                  f"source_pos_after={b['pos_after']}")
        else:
            a("borrow total=0")
        a(f"weighted rate={rate:.1f}% numerator={fmt_num(w_done)} "
          f"denominator={fmt_num(w_sum)} "
          f"weights=" + ",".join(f"{ch}:{fmt_num(weights.get(ch, 0))}"
                                 for ch in channels))
        a(f"summary frame={frame_total} sample={sample_total} "
          f"report={csv_rows} frame_layers_sum={sum_frame_layers} "
          f"drawn_layers_sum={sum_real_layers} unique_tids={unique_tids} "
          f"quota_sum={quota_sum} pairs={len(pairs)} "
          f"unresolved={len(unresolved)} "
          f"channel_order={','.join(channels)}")


if __name__ == "__main__":
    main()
