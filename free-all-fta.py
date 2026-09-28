#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Markov-only allocative FTA updater with top-down safety-target propagation.
Usage: python3 excel_to_pfta_alloc_v2.py workbook.xlsx --alloc
Output: workbook_alloc.xlsx unless --alloc-out is used.
"""
from __future__ import annotations
import argparse, re, math
from pathlib import Path
from collections import defaultdict, OrderedDict
from typing import Any, Dict, List, Optional, Set, Tuple
import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.worksheet.table import TableColumn
from openpyxl.worksheet.filters import AutoFilter

ID_SPLIT = " - "
TREE_COL_RE = re.compile(r"^L(\d{1,2})_(ID|TYPE)$", re.I)
GREEN_FILL = PatternFill(fill_type="solid", start_color="C6EFCE", end_color="C6EFCE")
RED_FILL = PatternFill(fill_type="solid", start_color="FFC7CE", end_color="FFC7CE")
YELLOW_FILL = PatternFill(fill_type="solid", start_color="FFF2CC", end_color="FFF2CC")


def clean_id(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).replace("\r", " ").replace("\n", " ").replace("\t", " ").strip()
    return s.split(ID_SPLIT, 1)[0].strip() if ID_SPLIT in s else s


def hkey(v: Any) -> str:
    return str(v).strip().lower() if v is not None else ""


def headers(ws) -> Dict[str, int]:
    return {hkey(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value is not None}


def ensure_col(ws, name: str, after: Optional[str] = None) -> int:
    hs = headers(ws)
    k = name.lower()
    if k in hs:
        return hs[k]
    if after and after.lower() in hs:
        c = hs[after.lower()] + 1
        ws.insert_cols(c)
        ws.cell(1, c).value = name
        return c
    c = ws.max_column + 1
    ws.cell(1, c).value = name
    return c


def rows_by_id(ws) -> Tuple[Dict[str, int], Dict[str, int]]:
    hs = headers(ws)
    idc = hs.get("id")
    out: Dict[str, int] = {}
    if idc:
        for r in range(2, ws.max_row + 1):
            rid = clean_id(ws.cell(r, idc).value)
            if rid:
                ws.cell(r, idc).value = rid
                out[rid] = r
    return hs, out


def ref(sheet: str, row: int, col: int, absref: bool = True) -> str:
    addr = f"${get_column_letter(col)}${row}" if absref else f"{get_column_letter(col)}{row}"
    return f"'{sheet.replace(chr(39), chr(39) * 2)}'!{addr}"


def plus(parts: List[str]) -> str:
    ps = [p for p in parts if p and p != "0"]
    return "(" + "+".join(ps) + ")" if ps else "0"


def prod(parts: List[str]) -> str:
    ps: List[str] = []
    for p in parts:
        if not p or p == "1":
            continue
        if p == "0":
            return "0"
        ps.append(p)
    return "(" + "*".join(ps) + ")" if ps else "1"


def model(v: Any) -> str:
    m = str(v).strip().lower().replace(" ", "") if v is not None else ""
    if m in ("constantrate", "constant_rate"):
        return "ConstantRate"
    if m == "fixed":
        return "Fixed"
    if m in ("probability", "fixed(probability)"):
        return "Probability"
    if m in ("multiplicity", "multiplicity(alloc)"):
        return "Multiplicity"
    if m == "true":
        return "True"
    if m == "false":
        return "False"
    return str(v).strip() if v is not None else "Unknown"


def tree_cols(ws) -> List[Tuple[int, int, Optional[int]]]:
    d: Dict[int, Dict[str, int]] = defaultdict(dict)
    for c in range(1, ws.max_column + 1):
        m = TREE_COL_RE.match(str(ws.cell(1, c).value or "").strip())
        if m:
            d[int(m.group(1))][m.group(2).upper()] = c
    return [(lvl, d[lvl]["ID"], d[lvl].get("TYPE")) for lvl in sorted(d) if "ID" in d[lvl]]


def parse_tree(ws, event_ids: Set[str], gate_ids: Set[str]):
    """Parse a vertically indented tree: one node per row, level given by Lxx_ID."""
    children = defaultdict(OrderedDict)
    gtype: Dict[str, str] = {}
    last_at_level: Dict[int, str] = {}
    cols = tree_cols(ws)
    for r in range(2, ws.max_row + 1):
        current = None
        for lvl, idc, typec in cols:
            rid = clean_id(ws.cell(r, idc).value)
            if not rid:
                continue
            ws.cell(r, idc).value = rid
            raw = str(ws.cell(r, typec).value).strip() if typec and ws.cell(r, typec).value is not None else ""
            typ = raw.upper() if raw else None
            current = (lvl, rid, typ)
            break
        if current is None:
            continue
        lvl, rid, typ = current
        if rid in gate_ids and typ and typ != "NULL":
            if typ not in ("AND", "OR"):
                raise ValueError(f"Unsupported gate type '{typ}' for gate '{rid}'. Only AND/OR are accepted.")
            gtype[rid] = typ
        parent = None
        for plvl in sorted([x for x in last_at_level if x < lvl], reverse=True):
            parent = last_at_level[plvl]
            break
        if parent and parent in gate_ids:
            children[parent][rid] = None
        last_at_level[lvl] = rid
        for dlvl in [x for x in list(last_at_level) if x > lvl]:
            del last_at_level[dlvl]
    return {k: list(v.keys()) for k, v in children.items()}, gtype


def update_table(ws, tname: str):
    """Safely resize an Excel structured table after inserted columns/rows.

    Excel is stricter than openpyxl: table ref, autoFilter ref, and the
    tableColumns collection must all describe exactly the same range.
    If the table ref is expanded to A:O while tableColumns still contains
    only 14 columns, Excel reports the workbook as needing repair and may
    fail to repair it.
    """
    if tname not in ws.tables:
        return

    tab = ws.tables[tname]
    minc, minr, _maxc, _maxr = range_boundaries(tab.ref)

    # Last data row: use the first column of the table, not ws.max_row,
    # because formulas/styles below the table may inflate ws.max_row.
    last = minr
    for r in range(ws.max_row, minr, -1):
        if ws.cell(r, minc).value not in (None, ""):
            last = r
            break

    # Last header column: extend to the last non-empty header in the table
    # header row. This covers columns inserted in the middle of a table.
    lastc = minc
    for c in range(ws.max_column, minc - 1, -1):
        if ws.cell(minr, c).value not in (None, ""):
            lastc = c
            break

    new_ref = f"{get_column_letter(minc)}{minr}:{get_column_letter(lastc)}{last}"
    tab.ref = new_ref

    # Keep autoFilter aligned with the table ref. A stale autoFilter ref is
    # another common cause of Excel repair messages.
    if getattr(tab, "autoFilter", None) is None:
        tab.autoFilter = AutoFilter(ref=new_ref)
    else:
        tab.autoFilter.ref = new_ref

    # Rebuild tableColumns to match the header row and new_ref exactly.
    # This is the critical fix for files that openpyxl can save but Excel
    # flags as corrupt/unrepairable.
    cols = []
    seen = set()
    for idx, c in enumerate(range(minc, lastc + 1), start=1):
        name = ws.cell(minr, c).value
        name = str(name).strip() if name not in (None, "") else f"Column{idx}"
        # Excel table column names must be unique.
        base = name
        n = 2
        while name in seen:
            name = f"{base}_{n}"
            n += 1
        seen.add(name)
        cols.append(TableColumn(id=idx, name=name))
    tab.tableColumns = cols


def build_markov_workbook(xlsx: str, out_xlsx: Optional[str] = None) -> str:
    src = Path(xlsx).expanduser().resolve()
    dst = Path(out_xlsx).expanduser().resolve() if out_xlsx else src.with_name(src.stem + "_alloc" + src.suffix)

    wb = openpyxl.load_workbook(src, data_only=False)
    we, wg, wt = wb["events"], wb["gates"], wb["tree"]

    ensure_col(wg, "calculated frequency", after="safety target")
    ensure_col(wg, "calculated mean repair time", after="calculated frequency")
    ensure_col(wg, "differences", after="calculated mean repair time")
    ensure_col(wg, "Formula")

    eh, erows = rows_by_id(we)
    gh, grows = rows_by_id(wg)
    eh, gh = headers(we), headers(wg)

    for c in ("model_type", "allocated value", "mean_repair_time"):
        if c not in eh:
            raise ValueError(f"Missing events column: {c}")
    for c in ("formula", "safety target", "calculated frequency", "calculated mean repair time", "differences"):
        if c not in gh:
            raise ValueError(f"Missing gates column: {c}")

    children, gtypes = parse_tree(wt, set(erows), set(grows))
    em, ea, et = eh["model_type"], eh["allocated value"], eh["mean_repair_time"]
    gf, gt, gc, gm, gd = gh["formula"], gh["safety target"], gh["calculated frequency"], gh["calculated mean repair time"], gh["differences"]

    def propagate_safety_targets_top_down():
        """Top-down allocation through AND-only chains.

        If a parent AND gate has exactly one child gate, allocate the child gate
        safety target by DIVIDING the parent safety target by the product of ALL
        event child allocated values:

            child target = parent safety target / product(event allocated values)

        This fixes the previous incorrect behaviour where some event values were
        multiplied with the parent target.

        Example:
            =(gates!$F$2*(events!$M$2*events!$M$5*events!$M$6*events!$M$7))
        becomes:
            =(gates!$F$2/(events!$M$2*events!$M$5*events!$M$6*events!$M$7))

        Propagation stops at OR gates, unknown children, or more/less than one
        child gate.
        """
        visited: Set[str] = set()

        def walk(parent_gid: str):
            if parent_gid in visited:
                return
            visited.add(parent_gid)
            if gtypes.get(parent_gid) != "AND":
                return
            chs = children.get(parent_gid, [])
            gate_chs = [ch for ch in chs if ch in grows]
            event_chs = [ch for ch in chs if ch in erows]
            unknown_chs = [ch for ch in chs if ch not in grows and ch not in erows]
            if unknown_chs or len(gate_chs) != 1:
                return

            child_gid = gate_chs[0]
            parent_row = grows[parent_gid]
            child_row = grows[child_gid]

            # Correct rule: all event child values are denominator factors.
            # No event value is multiplied with the parent safety target.
            denominator: List[str] = []
            for eid in event_chs:
                er = erows[eid]
                denominator.append(ref("events", er, ea))

            parent_target_ref = ref("gates", parent_row, gt)
            den_expr = prod(denominator)
            target_expr = parent_target_ref if den_expr == "1" else f"({parent_target_ref}/{den_expr})"
            wg.cell(child_row, gt).value = "=" + target_expr
            wg.cell(child_row, gt).fill = YELLOW_FILL
            walk(child_gid)

        # Start from every gate that already has a target. This supports AND sub-branches
        # even when the top of the tree is OR or has multiple child gates.
        for gid in grows:
            if gid in children and wg.cell(grows[gid], gt).value not in (None, ""):
                walk(gid)

    propagate_safety_targets_top_down()

    visiting: Set[str] = set()
    done: Set[str] = set()

    def evop(eid: str):
        r = erows[eid]
        m = model(we.cell(r, em).value)
        val = ref("events", r, ea)
        T = ref("events", r, et)
        if m in ("ConstantRate", "Fixed"):
            return {"id": eid, "kind": "rate", "l": val, "T": T, "m": m}
        if m == "Probability":
            return {"id": eid, "kind": "factor", "f": val, "m": m}
        if m == "Multiplicity":
            return {"id": eid, "kind": "factor", "f": val, "m": m}
        if m == "True":
            return {"id": eid, "kind": "factor", "f": "1", "m": m}
        if m == "False":
            return {"id": eid, "kind": "factor", "f": "0", "m": m}
        return {"id": eid, "kind": "rate", "l": val, "T": T, "m": m}

    def gop(gid: str):
        compute(gid)
        r = grows[gid]
        return {"id": gid, "kind": "rate", "l": ref("gates", r, gc), "T": ref("gates", r, gm), "m": "Gate"}

    def and_formula(ops):
        rates = [o for o in ops if o["kind"] == "rate"]
        factors = [o for o in ops if o["kind"] == "factor"]
        fexpr = prod([o["f"] for o in factors])
        if not rates:
            lexpr, texpr = fexpr, ""
        elif len(rates) == 1:
            lexpr, texpr = prod([fexpr, rates[0]["l"]]), rates[0]["T"]
        else:
            lt = [f"({o['l']}*{o['T']})" for o in rates]
            inv = [f"(1/{o['T']})" for o in rates]
            lexpr = prod([fexpr, prod(lt), plus(inv)])
            texpr = f"(1/{plus(inv)})"
        text = "AND Markov: λEQ = "
        if factors:
            text += " × ".join(o["id"] for o in factors) + " × "
        text += ("∏(λi×Ti) × Σ(1/Ti), rate children: " + ", ".join(o["id"] for o in rates)) if len(rates) > 1 else (" × ".join(o["id"] for o in rates) if rates else "factor-only product")
        if rates:
            text += "; TEQ = 1 / Σ(1/Ti)"
        return "=" + lexpr, ("=" + texpr if texpr else ""), text

    def or_formula(ops):
        rates = [o for o in ops if o["kind"] == "rate"]
        factors = [o for o in ops if o["kind"] == "factor"]
        lexpr = plus([o["l"] for o in rates] + [o["f"] for o in factors])
        texpr = ""
        if rates:
            texpr = f"({plus([f'({o["l"]}*{o["T"]})' for o in rates])}/{plus([o['l'] for o in rates])})"
        text = "OR: λEQ = Σλi over children: " + ", ".join(o["id"] for o in ops)
        if rates:
            text += "; TEQ = Σ(λi×Ti) / Σλi"
        if factors:
            text += "; warning: factor-only OR children have no TEQ contribution"
        return "=" + lexpr, ("=" + texpr if texpr else ""), text

    def compute(gid: str):
        if gid in done:
            return
        if gid in visiting:
            raise ValueError(f"Cycle detected involving gate {gid}")
        visiting.add(gid)
        chs = children.get(gid, [])
        r = grows[gid]
        if chs:
            typ = gtypes.get(gid)
            if typ not in ("AND", "OR"):
                raise ValueError(f"Gate {gid} has children but no AND/OR type")
            ops = []
            for ch in chs:
                if ch in erows:
                    ops.append(evop(ch))
                elif ch in grows:
                    ops.append(gop(ch))
                else:
                    raise ValueError(f"Child {ch} under {gid} not found in events/gates")
            lf, tf, txt = and_formula(ops) if typ == "AND" else or_formula(ops)
            wg.cell(r, gf).value = txt
            wg.cell(r, gc).value = lf
            wg.cell(r, gm).value = tf
            wg.cell(r, gd).value = f'=IF({ref("gates", r, gt, False)}="","",{ref("gates", r, gt, False)}-{ref("gates", r, gc, False)})'
        visiting.remove(gid)
        done.add(gid)

    for gid in list(grows):
        if gid in children:
            compute(gid)

    dr = f"{get_column_letter(gd)}2:{get_column_letter(gd)}{wg.max_row}"
    wg.conditional_formatting.add(dr, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=GREEN_FILL))
    wg.conditional_formatting.add(dr, CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL))
    for gid in children:
        if gid in grows:
            r = grows[gid]
            for c in (gf, gc, gm):
                wg.cell(r, c).fill = YELLOW_FILL

    update_table(we, "tbl_events")
    update_table(wg, "tbl_gates")
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass
    wb.save(dst)
    openpyxl.load_workbook(dst, data_only=False).close()
    return str(dst)





def build_pure_workbook(xlsx: str, out_xlsx: Optional[str] = None) -> str:
    """Pure allocative FTA updater.

    This mode does not use Markov formulas. It calculates probabilities with:
      - AND: qEQ = product(qi)
      - OR : qEQ = sum(qi)
      - For ConstantRate/Fixed events: qi = 1 - EXP(-(lambda_i * T_i))
      - Equivalent intensity/frequency: wEQ = qEQ / TEQ

    The existing --alloc Markov workflow is intentionally untouched.
    """
    src = Path(xlsx).expanduser().resolve()
    dst = Path(out_xlsx).expanduser().resolve() if out_xlsx else src.with_name(src.stem + "_pure" + src.suffix)

    wb = openpyxl.load_workbook(src, data_only=False)
    we, wg, wt = wb["events"], wb["gates"], wb["tree"]

    ensure_col(wg, "calculated frequency", after="safety target")
    ensure_col(wg, "calculated mean repair time", after="calculated frequency")
    ensure_col(wg, "differences", after="calculated mean repair time")
    ensure_col(wg, "calculated probability", after="differences")
    ensure_col(wg, "Formula")

    eh, erows = rows_by_id(we)
    gh, grows = rows_by_id(wg)
    eh, gh = headers(we), headers(wg)

    for c in ("model_type", "allocated value", "mean_repair_time"):
        if c not in eh:
            raise ValueError(f"Missing events column: {c}")
    for c in ("formula", "calculated frequency", "calculated mean repair time", "calculated probability"):
        if c not in gh:
            raise ValueError(f"Missing gates column: {c}")

    children, gtypes = parse_tree(wt, set(erows), set(grows))
    em, ea, et = eh["model_type"], eh["allocated value"], eh["mean_repair_time"]
    ep = eh.get("probability")
    gf, gc, gm, gp = gh["formula"], gh["calculated frequency"], gh["calculated mean repair time"], gh["calculated probability"]
    gd, gt = gh.get("differences"), gh.get("safety target")

    visiting: Set[str] = set()
    done: Set[str] = set()

    def evop(eid: str):
        r = erows[eid]
        m = model(we.cell(r, em).value)
        lam = ref("events", r, ea)
        T = ref("events", r, et)
        if m in ("ConstantRate", "Fixed"):
            # Pure probability of at least one occurrence during T.
            q = f"(1-EXP(-({lam}*{T})))"
            return {"id": eid, "q": q, "raw_q": q, "w": lam, "T": T, "m": m}
        if m == "Probability":
            q = ref("events", r, ep) if ep else lam
            return {"id": eid, "q": q, "raw_q": q, "w": "", "T": "", "m": m}
        if m == "Multiplicity":
            # Multiplicity is a factor, not a probability, but it participates
            # in the raw probability expression for the parent gate.
            return {"id": eid, "q": lam, "raw_q": lam, "w": "", "T": "", "m": m}
        if m == "True":
            return {"id": eid, "q": "1", "raw_q": "1", "w": "", "T": "", "m": m}
        if m == "False":
            return {"id": eid, "q": "0", "raw_q": "0", "w": "", "T": "", "m": m}
        q = f"(1-EXP(-({lam}*{T})))"
        return {"id": eid, "q": q, "raw_q": q, "w": lam, "T": T, "m": m}

    def gop(gid: str):
        compute(gid)
        r = grows[gid]
        w_ref = ref("gates", r, gc)
        t_ref = ref("gates", r, gm)
        # For parent w calculations, do not reuse the capped q of a child gate.
        # Reconstruct the child raw probability as w_child * T_child.
        return {
            "id": gid,
            "q": ref("gates", r, gp),
            "raw_q": f"({w_ref}*{t_ref})",
            "w": w_ref,
            "T": t_ref,
            "m": "Gate",
        }

    def max_expr(parts: List[str]) -> str:
        ps = [p for p in parts if p]
        if not ps:
            return ""
        if len(ps) == 1:
            return ps[0]
        return "MAX(" + ",".join(ps) + ")"

    def pure_formula(typ: str, ops):
        qparts = [o["q"] for o in ops if o.get("q") not in (None, "")]
        raw_qparts = [o["raw_q"] for o in ops if o.get("raw_q") not in (None, "")]
        wparts = [o["w"] for o in ops if o.get("w") not in (None, "")]
        tparts = [o["T"] for o in ops if o.get("T") not in (None, "")]
        texpr = max_expr(tparts)

        if typ == "AND":
            q_raw_for_q = prod(qparts)
            raw_for_w = prod(raw_qparts)
            qexpr = f"MIN(1,{q_raw_for_q})"
            wexpr = f"({raw_for_w}/{texpr})" if texpr else ""
            text = (
                "PURE AND: rate-event qi = 1-EXP(-(lambda_i*T_i)); "
                "qEQ = MIN(1, product(qi)); "
                "wEQ = raw_product(qi) / T before MIN saturation; "
                "for a child gate raw_q_child = w_child*T_child; T = MAX(Ti); children: "
            )
        else:
            q_raw_for_q = plus(qparts)
            raw_for_w = plus(raw_qparts)
            qexpr = f"MIN(1,{q_raw_for_q})"
            # For OR, frequency/intensity is additive. This is the key
            # correction compared with raw_q/T when one branch is already capped.
            wexpr = plus(wparts) if wparts else (f"({raw_for_w}/{texpr})" if texpr else "")
            text = (
                "PURE OR: rate-event qi = 1-EXP(-(lambda_i*T_i)); "
                "qEQ = MIN(1, sum(qi)); "
                "wEQ = sum(wi) for rate/gate branches, not capped_q/T; "
                "T = MAX(Ti); children: "
            )
        text += ", ".join(o["id"] for o in ops)
        return qexpr, texpr, wexpr, text

    def compute(gid: str):
        if gid in done:
            return
        if gid in visiting:
            raise ValueError(f"Cycle detected involving gate {gid}")
        visiting.add(gid)
        chs = children.get(gid, [])
        r = grows[gid]
        if chs:
            typ = gtypes.get(gid)
            if typ not in ("AND", "OR"):
                raise ValueError(f"Gate {gid} has children but no AND/OR type")
            ops = []
            for ch in chs:
                if ch in erows:
                    ops.append(evop(ch))
                elif ch in grows:
                    ops.append(gop(ch))
                else:
                    raise ValueError(f"Child {ch} under {gid} not found in events/gates")
            qexpr, texpr, wexpr, txt = pure_formula(typ, ops)
            wg.cell(r, gf).value = txt
            wg.cell(r, gp).value = "=" + qexpr
            wg.cell(r, gm).value = "=" + texpr if texpr else ""
            # In --pure, q is capped with MIN(1,...), but w must be derived
            # from the raw/uncapped probability expression. Otherwise a branch
            # with multiplicity can saturate q to 1 and incorrectly give w=1/T.
            wg.cell(r, gc).value = "=" + wexpr if wexpr else ""
            if gd and gt:
                wg.cell(r, gd).value = f'=IF({ref("gates", r, gt, False)}="","",{ref("gates", r, gt, False)}-{ref("gates", r, gc, False)})'
        visiting.remove(gid)
        done.add(gid)

    for gid in list(grows):
        if gid in children:
            compute(gid)

    if gd:
        dr = f"{get_column_letter(gd)}2:{get_column_letter(gd)}{wg.max_row}"
        wg.conditional_formatting.add(dr, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=GREEN_FILL))
        wg.conditional_formatting.add(dr, CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL))

    for ws, tn in ((we, "tbl_events"), (wg, "tbl_gates")):
        update_table(ws, tn)
    wb.save(dst)
    return str(dst)

# --- Native SVG FTA renderer v7 (no PFTA model, no Graphviz) -----------------
# Fixes: display evaluated values only (never formulas), same width for gates and
# AND/OR operators, same height for AND/OR operators, larger fonts.
import os, html as _html, shutil as _shutil, ast as _ast, operator as _operator
from openpyxl.utils.cell import coordinate_to_tuple as _coord_to_tuple

_FILL = '#ffffdd'
_STROKE = '#000000'
_FONT = 'Courier New, Courier, monospace'
_GATE_W = 245
_TITLE_H = 76
_ID_H = 26
# PFTA-like operator: compact symbol + value box, not a huge distorted gate-wide arc
_OP_W = _GATE_W
_OP_H = 185
_EVENT_VALUE_H = 84
_LEVEL_GAP = 128
_NODE_GAP = 46

_BIN = {
    _ast.Add: _operator.add,
    _ast.Sub: _operator.sub,
    _ast.Mult: _operator.mul,
    _ast.Div: _operator.truediv,
    _ast.Pow: _operator.pow,
    _ast.BitXor: _operator.pow,
}
_UN = {_ast.UAdd:lambda x:x, _ast.USub:lambda x:-x}

def _num(v):
    if v is None or v == '': return None
    if isinstance(v, bool): return float(v)
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip()
    if not s or s.startswith('='): return None
    try: return float(s.replace(',', '.'))
    except Exception: return None

def _fmt(v, unit=''):
    n = _num(v)
    if n is None:
        return ''   # IMPORTANT: never display formulas or non-evaluated strings as values
    txt = f'{n:.3E}' if abs(n) != 0 and (abs(n) < 1e-3 or abs(n) >= 1e4) else f'{n:.3g}'
    return txt + unit

def _wrap(txt, width=30, max_lines=4):
    txt = '' if txt is None else str(txt).replace('\n',' ').replace('\r',' ').strip()
    if not txt: return []
    words, lines, cur = txt.split(), [], ''
    for w in words:
        if len(cur) + len(w) + (1 if cur else 0) <= width:
            cur = (cur + ' ' + w).strip()
        else:
            if cur: lines.append(cur)
            cur = w
        if len(lines) >= max_lines: break
    if cur and len(lines) < max_lines: lines.append(cur)
    return lines[:max_lines]

def _safe_eval(expr):
    def ev(n):
        if isinstance(n, _ast.Expression): return ev(n.body)
        if isinstance(n, _ast.Constant) and isinstance(n.value, (int, float)): return float(n.value)
        if hasattr(_ast, 'Num') and isinstance(n, _ast.Num): return float(n.n)
        if isinstance(n, _ast.BinOp) and type(n.op) in _BIN: return _BIN[type(n.op)](ev(n.left), ev(n.right))
        if isinstance(n, _ast.UnaryOp) and type(n.op) in _UN: return _UN[type(n.op)](ev(n.operand))
        if isinstance(n, _ast.Call) and isinstance(n.func, _ast.Name):
            name = n.func.id.upper()
            args = [ev(a) for a in n.args]
            if name == 'MIN': return min(args)
            if name == 'MAX': return max(args)
            if name == 'SUM': return sum(args)
            if name == 'PRODUCT':
                out = 1.0
                for a in args: out *= a
                return out
            if name == 'EXP' and len(args) == 1: return math.exp(args[0])
            if name == 'ABS' and len(args) == 1: return abs(args[0])
            if name == 'POWER' and len(args) == 2: return args[0] ** args[1]
            if name == 'SQRT' and len(args) == 1: return math.sqrt(args[0])
            if name == 'ROUND' and len(args) == 2: return round(args[0], int(args[1]))
        raise ValueError(type(n).__name__)
    return ev(_ast.parse(expr, mode='eval'))

def _defined_name_value(wbf, wbv, name, memo=None):
    """Return the scalar value of an Excel defined name, if it resolves to a cell/range.

    Supports workbook-level and sheet-scoped names where openpyxl exposes the
    destination as (sheet_name, cell_or_range). If the name points to a range,
    the top-left cell is used, which is the expected behaviour for scalar
    named ranges such as sdt -> specificData!$C$7.
    """
    if not name:
        return None
    candidates = {name, name.strip(), name.strip().lower(), name.strip().upper()}
    for cand in candidates:
        try:
            dn = wbf.defined_names.get(cand)
        except Exception:
            dn = None
        if dn is None:
            continue
        try:
            destinations = list(dn.destinations)
        except Exception:
            destinations = []
        for sh, ref_text in destinations:
            sh = str(sh).strip("'")
            if sh not in wbf.sheetnames:
                continue
            # If the defined name is a range, take the first/top-left cell.
            first_ref = str(ref_text).split(':', 1)[0].replace('$', '')
            try:
                r2, c2 = _coord_to_tuple(first_ref)
            except Exception:
                continue
            return _eval_cell(wbf, wbv, sh, r2, c2, memo)
    return None


def _replace_defined_names_in_expr(wbf, wbv, expr, memo=None):
    """Replace scalar Excel defined names by their numeric values in an expression.

    This is intentionally called after worksheet/cell references and structured
    references have already been expanded, so it does not interfere with sheet
    names before '!'. It also avoids replacing known function names.
    """
    if not isinstance(expr, str) or not expr:
        return expr
    function_names = {
        'IF', 'IFERROR', 'VALUE', 'SUBSTITUTE', 'SUM', 'PRODUCT', 'MIN', 'MAX',
        'AND', 'OR', 'NOT', 'TRUE', 'FALSE', 'ABS', 'ROUND', 'POWER', 'SQRT', 'EXP'
    }

    def repl(m):
        token = m.group(0)
        if token.upper() in function_names:
            return token
        # Do not replace tokens that are part of sheet references like Sheet1!A1.
        end = m.end()
        if end < len(expr) and expr[end:end+1] == '!':
            return token
        val = _defined_name_value(wbf, wbv, token, memo)
        n = _num(val)
        if n is None:
            return token
        return str(n)

    return re.sub(r"\b[A-Za-z_][A-Za-z0-9_.]*\b", repl, expr)


def _eval_cell(wbf, wbv, sheet, row, col, memo=None):
    if col is None: return None
    if memo is None: memo = {}
    key = (sheet, row, col)
    if key in memo: return memo[key]
    cached = wbv[sheet].cell(row, col).value
    if hasattr(cached, 'text'):
        cached = None
    if cached not in (None, ''):
        memo[key] = cached; return cached
    ws = wbf[sheet]
    raw = ws.cell(row, col).value
    if hasattr(raw, 'text'):
        raw = raw.text
    if raw is None or raw == '' or not (isinstance(raw, str) and raw.startswith('=')):
        memo[key] = raw; return raw
    expr = raw[1:]

    # Handle the template MTTR formula before replacing structured references.
    # Otherwise VALUE/SUBSTITUTE(...) would remain in the expression and fail.
    if expr.upper().startswith('IFERROR(') and 'detection_time' in expr.lower() and 'negation_time' in expr.lower():
        hs = headers(ws)
        det = _eval_cell(wbf, wbv, sheet, row, hs.get('detection_time', col), memo)
        neg = _eval_cell(wbf, wbv, sheet, row, hs.get('negation_time', col), memo)
        memo[key] = (_num(det) or 0.0) + (_num(neg) or 0.0)
        return memo[key]

    def repl_struct(m):
        c = headers(ws).get(m.group(1).strip().lower())
        val = _eval_cell(wbf, wbv, sheet, row, c, memo) if c else 0
        return str(_num(val) or 0)
    expr = re.sub(r"tbl_events\[\[#This Row\],\[([^\]]+)\]\]", repl_struct, expr, flags=re.I)

    def repl_quoted(m):
        sh = m.group(1); coord = (m.group(2)+m.group(3)).replace('$','')
        if sh not in wbf.sheetnames: return '0'
        r2, c2 = _coord_to_tuple(coord)
        return str(_num(_eval_cell(wbf, wbv, sh, r2, c2, memo)) or 0)
    expr = re.sub(r"'([^']+)'!\$?([A-Z]+)\$?(\d+)", repl_quoted, expr)

    def repl_unquoted(m):
        sh = m.group(1); coord = (m.group(2)+m.group(3)).replace('$','')
        if sh not in wbf.sheetnames: return m.group(0)
        r2, c2 = _coord_to_tuple(coord)
        return str(_num(_eval_cell(wbf, wbv, sh, r2, c2, memo)) or 0)
    expr = re.sub(r"([A-Za-z_][A-Za-z0-9_]*)!\$?([A-Z]+)\$?(\d+)", repl_unquoted, expr)

    def repl_local(m):
        r2, c2 = _coord_to_tuple(m.group(1).replace('$',''))
        return str(_num(_eval_cell(wbf, wbv, sheet, r2, c2, memo)) or 0)
    expr = re.sub(r"(?<![A-Za-z0-9_!])\$?([A-Z]+\$?\d+)", repl_local, expr)

    # Resolve scalar workbook/sheet named ranges wherever they appear, e.g.
    # =sdt or formulas depending on a cell containing =sdt.
    expr = _replace_defined_names_in_expr(wbf, wbv, expr, memo)

    try: val = _safe_eval(expr)
    except Exception: val = None  # critical: do not display raw formula
    memo[key] = val
    return val

def _truthy(v):
    if isinstance(v, bool): return v
    return False if v is None else str(v).strip().lower() in ('true','yes','y','1','x')

def _diagram_data(xlsx, pure=False):
    wbf = openpyxl.load_workbook(xlsx, data_only=False)
    wbv = openpyxl.load_workbook(xlsx, data_only=True)
    we, wg, wt = wbf['events'], wbf['gates'], wbf['tree']
    eh, erows = rows_by_id(we); gh, grows = rows_by_id(wg)
    eh, gh = headers(we), headers(wg)
    children, gtypes = parse_tree(wt, set(erows), set(grows))
    memo, events, gates = {}, {}, {}
    for eid, r in erows.items():
        mt = model(we.cell(r, eh.get('model_type')).value) if eh.get('model_type') else 'Unknown'
        events[eid] = {
            'label': we.cell(r, eh.get('label')).value if eh.get('label') else eid,
            'model': mt,
            'alloc': _eval_cell(wbf, wbv, 'events', r, eh.get('allocated value'), memo),
            'mrt': _eval_cell(wbf, wbv, 'events', r, eh.get('mean_repair_time'), memo),
            'prob': _eval_cell(wbf, wbv, 'events', r, eh.get('probability'), memo) if eh.get('probability') else None,
        }
    page_col = gh.get('is_paged') or gh.get('is_page')
    for gid, r in grows.items():
        gates[gid] = {
            'label': wg.cell(r, gh.get('label')).value if gh.get('label') else gid,
            'w': _eval_cell(wbf, wbv, 'gates', r, gh.get('calculated frequency'), memo),
            'sdt': _eval_cell(wbf, wbv, 'gates', r, gh.get('calculated mean repair time'), memo),
            'q': _eval_cell(wbf, wbv, 'gates', r, gh.get('calculated probability'), memo) if gh.get('calculated probability') else None,
            'is_page': _truthy(wg.cell(r, page_col).value) if page_col else False,
        }
    # Fallback/recalculation for diagram values.
    # The SVG shall show values, never Excel formulas. If Excel has no cached
    # result, compute display values bottom-up from child events/gates.
    # This specifically covers AND gates whose children are events: w is taken
    # from events[Allocated value] and sdt from events[mean_repair_time].
    computed = {}

    def as_num(v):
        return _num(v)

    def event_operand(eid):
        e = events[eid]
        mt = e.get('model')
        if mt in ('ConstantRate', 'Fixed'):
            return {'kind': 'rate', 'w': as_num(e.get('alloc')), 'sdt': as_num(e.get('mrt'))}
        if mt == 'Probability':
            return {'kind': 'factor', 'factor': as_num(e.get('alloc'))}
        if mt == 'Multiplicity':
            return {'kind': 'factor', 'factor': as_num(e.get('alloc'))}
        if mt == 'True':
            return {'kind': 'factor', 'factor': 1.0}
        if mt == 'False':
            return {'kind': 'factor', 'factor': 0.0}
        return {'kind': 'rate', 'w': as_num(e.get('alloc')), 'sdt': as_num(e.get('mrt'))}

    def gate_operand(gid):
        w, sdt = compute_gate_value(gid)
        return {'kind': 'rate', 'w': w, 'sdt': sdt}

    def compute_gate_value(gid):
        if gid in computed:
            return computed[gid]
        ops = []
        for ch in children.get(gid, []):
            if ch in erows:
                ops.append(event_operand(ch))
            elif ch in grows:
                ops.append(gate_operand(ch))
        typ = gtypes.get(gid)
        rates = [o for o in ops if o.get('kind') == 'rate' and o.get('w') is not None]
        factors = [o.get('factor') for o in ops if o.get('kind') == 'factor' and o.get('factor') is not None]
        factor = 1.0
        for f in factors:
            factor *= f
        w_val = None
        sdt_val = None
        if typ == 'AND':
            usable_rates = [o for o in rates if o.get('sdt') not in (None, 0)]
            if not usable_rates:
                w_val = factor if factors else None
            elif len(usable_rates) == 1:
                w_val = factor * usable_rates[0]['w']
                sdt_val = usable_rates[0]['sdt']
            else:
                prod_lt = 1.0
                sum_inv_t = 0.0
                for o in usable_rates:
                    prod_lt *= o['w'] * o['sdt']
                    sum_inv_t += 1.0 / o['sdt']
                w_val = factor * prod_lt * sum_inv_t
                sdt_val = 1.0 / sum_inv_t if sum_inv_t else None
        elif typ == 'OR':
            total_w = sum(o['w'] for o in rates if o.get('w') is not None)
            w_val = total_w
            if total_w:
                num = sum(o['w'] * o['sdt'] for o in rates if o.get('w') is not None and o.get('sdt') is not None)
                sdt_val = num / total_w if num else None

        # Prefer workbook numeric values when they exist. If a cell contains a
        # formula string or no cached value, _num returns None and fallback is used.
        existing_w = as_num(gates.get(gid, {}).get('w'))
        existing_sdt = as_num(gates.get(gid, {}).get('sdt'))
        if existing_w is not None:
            w_val = existing_w
        if existing_sdt is not None:
            sdt_val = existing_sdt
        computed[gid] = (w_val, sdt_val)
        return computed[gid]

    # In Markov/--alloc diagrams, compute missing display values with the Markov
    # fallback. In --pure diagrams, do NOT run this Markov fallback because it can
    # overwrite the workbook values from gates[calculated frequency] and
    # gates[calculated mean repair time] with Markov-equivalent values.
    # For --pure, gate_pure() below first keeps the evaluated workbook values and
    # only computes q/w/t as a fallback when the workbook value is unavailable.
    if not pure:
        for gid in grows:
            if gid in children:
                w_val, sdt_val = compute_gate_value(gid)
                gates[gid]['w'] = w_val
                gates[gid]['sdt'] = sdt_val

    if pure:
        pure_done = {}
        def event_pure(eid):
            e = events[eid]
            mt = e.get('model')
            w = as_num(e.get('alloc'))
            t = as_num(e.get('mrt'))
            if mt == 'Probability':
                q = as_num(e.get('prob'))
                if q is None:
                    q = as_num(e.get('alloc'))
                return q, w, t, q
            if mt == 'Multiplicity':
                m = as_num(e.get('alloc'))
                return m, None, None, m
            if mt == 'True':
                return 1.0, None, None, 1.0
            if mt == 'False':
                return 0.0, None, None, 0.0
            q = 1.0 - math.exp(-(w * t)) if w is not None and t is not None else None
            return q, w, t, q

        def gate_pure(gid):
            if gid in pure_done:
                return pure_done[gid]
            qs, raw_qs, ws, ts = [], [], [], []
            for ch in children.get(gid, []):
                if ch in erows:
                    q, w, t, raw_q = event_pure(ch)
                elif ch in grows:
                    q, w, t, raw_q = gate_pure(ch)
                else:
                    continue
                if q is not None:
                    qs.append(q)
                if raw_q is not None:
                    raw_qs.append(raw_q)
                if w is not None:
                    ws.append(w)
                if t not in (None, 0):
                    ts.append(t)
            typ = gtypes.get(gid)
            raw_qv = None
            if typ == 'AND':
                qprod = 1.0
                for q in qs:
                    qprod *= q
                qv = min(1.0, qprod)
                raw_qv = 1.0
                for rq in raw_qs:
                    raw_qv *= rq
            elif typ == 'OR':
                qsum = sum(qs) if qs else None
                qv = min(1.0, qsum) if qsum is not None else None
                raw_qv = sum(raw_qs) if raw_qs else None
            else:
                qv = None
            tv = max(ts) if ts else None
            if typ == 'OR' and ws:
                wv = sum(ws)
            else:
                wv = (raw_qv / tv) if (raw_qv is not None and tv not in (None, 0)) else None
            # raw_q for a gate as seen by its parent: w*T, not capped q.
            gate_raw_q = (wv * tv) if (wv is not None and tv not in (None, 0)) else qv

            existing_q = as_num(gates[gid].get('q'))
            existing_w = as_num(gates[gid].get('w'))
            existing_t = as_num(gates[gid].get('sdt'))

            # In --pure diagrams, display the values from the pure workbook
            # gates sheet whenever available:
            #   calculated probability      -> q
            #   calculated frequency        -> w
            #   calculated mean repair time -> t
            # The recursive calculation below is only a fallback when a gate
            # cell has no cached/evaluable numeric value.
            gates[gid]['q'] = existing_q if existing_q is not None else qv
            gates[gid]['w'] = existing_w if existing_w is not None else wv
            gates[gid]['sdt'] = existing_t if existing_t is not None else tv

            sel_w = as_num(gates[gid].get('w'))
            sel_t = as_num(gates[gid].get('sdt'))
            selected_raw_q = (sel_w * sel_t) if (sel_w is not None and sel_t not in (None, 0)) else gate_raw_q
            pure_done[gid] = (
                as_num(gates[gid].get('q')),
                sel_w,
                sel_t,
                selected_raw_q,
            )
            return pure_done[gid]

        for gid in grows:
            if gid in children:
                gate_pure(gid)

    all_gate_children = {ch for chs in children.values() for ch in chs if ch in grows}
    roots = [gid for gid in grows if gid in children and gid not in all_gate_children] or ([next(iter(grows))] if grows else [])
    wbf.close(); wbv.close()
    return children, gtypes, erows, grows, events, gates, roots, pure

class FtaSvg:
    def __init__(self, data, page_files):
        if len(data) == 8:
            self.children, self.gtypes, self.erows, self.grows, self.events, self.gates, self.roots, self.pure = data
        else:
            self.children, self.gtypes, self.erows, self.grows, self.events, self.gates, self.roots = data
            self.pure = False
        self.page_files = page_files
        self.pos = {}
        self.svg = []

    def width_of(self, node, root):
        if node in self.grows:
            if self.gates.get(node,{}).get('is_page') and node != root: return _GATE_W + 30
            kids = self.children.get(node, [])
            if not kids: return _GATE_W + 30
            return max(_GATE_W + 30, sum(self.width_of(k, root) for k in kids) + _NODE_GAP*(len(kids)-1))
        return _GATE_W + 30

    def _child_path(self, parent_path, idx, child_id):
        return parent_path + ((idx, child_id),)

    def layout_gate(self, gid, x0, width, y, root, path=None):
        if path is None:
            path = ((0, gid),)
        cx = x0 + width/2
        self.pos[('gate', path)] = (cx-_GATE_W/2, y, _GATE_W, _TITLE_H+_ID_H+8)
        if gid in self.children:
            self.pos[('op', path)] = (cx-_OP_W/2, y+_TITLE_H+_ID_H+26, _OP_W, _OP_H)
            kids = self.children.get(gid, [])
            widths = [self.width_of(k, root) for k in kids]
            total = sum(widths) + _NODE_GAP*max(0,len(widths)-1)
            cur = cx - total/2
            child_y = y + _TITLE_H + _ID_H + 26 + _OP_H + _LEVEL_GAP
            for i,(k,w) in enumerate(zip(kids,widths)):
                child_path = self._child_path(path, i, k)
                if k in self.grows and not (self.gates.get(k,{}).get('is_page') and k != root):
                    self.layout_gate(k, cur, w, child_y, root, child_path)
                else:
                    self.pos[('leaf', child_path)] = (cur+w/2-_GATE_W/2, child_y, _GATE_W, _TITLE_H+_ID_H+_EVENT_VALUE_H+30)
                cur += w + _NODE_GAP

    def render(self, root, filepath):
        self.pos = {}
        total = self.width_of(root, root)
        root_path = ((0, root),)
        self.layout_gate(root, 40, total, 40, root, root_path)
        maxx = max(x+w for x,y,w,h in self.pos.values())+60; maxy = max(y+h for x,y,w,h in self.pos.values())+80
        self.svg = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{maxx:.0f}" height="{maxy:.0f}" viewBox="0 0 {maxx:.0f} {maxy:.0f}">',
                    f'<style>text{{font-family:{_FONT};font-size:14px}} .small{{font-size:11px}} .val{{font-size:13px;font-weight:bold}}</style>']
        self.draw_gate(root, root, root_path)
        self.svg.append('</svg>')
        Path(filepath).write_text('\n'.join(self.svg), encoding='utf-8')

    def esc(self,s): return _html.escape(str(s))
    def rect(self,x,y,w,h): self.svg.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" fill="{_FILL}" stroke="{_STROKE}" stroke-width="1.5"/>')
    def line(self,x1,y1,x2,y2): self.svg.append(f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" stroke="{_STROKE}" stroke-width="1.4"/>')
    def texts(self, lines, cx, y, cls=''):
        c = f' class="{cls}"' if cls else ''
        for i,l in enumerate(lines): self.svg.append(f'<text{c} x="{cx:.1f}" y="{y+i*16:.1f}" text-anchor="middle">{self.esc(l)}</text>')
    def _pure_lines_for_gate(self,gid):
        lines=[]
        q=_fmt(self.gates.get(gid,{}).get('q'))
        w=_fmt(self.gates.get(gid,{}).get('w'),'/h')
        t=_fmt(self.gates.get(gid,{}).get('sdt'),'h')
        if q: lines.append('q = '+q)
        if w: lines.append('w = '+w)
        if t: lines.append('t = '+t)
        return lines or [self.gtypes.get(gid,'GATE')]
    def op_lines(self,gid):
        if getattr(self,'pure',False):
            return self._pure_lines_for_gate(gid)
        lines=[]
        w=_fmt(self.gates.get(gid,{}).get('w'),'/h')
        s=_fmt(self.gates.get(gid,{}).get('sdt'),'h')
        if w: lines.append('w = '+w)
        if s: lines.append('sdt = '+s)
        return lines or [self.gtypes.get(gid,'GATE')]
    def transfer_lines(self,gid):
        lines=[gid]
        if getattr(self,'pure',False):
            lines += self._pure_lines_for_gate(gid)
            return lines
        w=_fmt(self.gates.get(gid,{}).get('w'),'/h')
        s=_fmt(self.gates.get(gid,{}).get('sdt'),'h')
        if w: lines.append('w = '+w)
        if s: lines.append('sdt = '+s)
        return lines
    def transfer_value_lines(self,gid):
        if getattr(self,'pure',False):
            return self._pure_lines_for_gate(gid)
        lines=[]
        w=_fmt(self.gates.get(gid,{}).get('w'),'/h')
        s=_fmt(self.gates.get(gid,{}).get('sdt'),'h')
        if w: lines.append('w = '+w)
        if s: lines.append('sdt = '+s)
        return lines
    def event_lines(self,eid):
        e=self.events.get(eid,{}); mt=e.get('model','Unknown')
        if getattr(self,'pure',False):
            if mt in ('ConstantRate','Fixed'):
                w0=_num(e.get('alloc')); t0=_num(e.get('mrt'))
                q=_fmt((1.0 - math.exp(-(w0*t0))) if w0 is not None and t0 is not None else None)
                w=_fmt(e.get('alloc'),'/h')
                t=_fmt(e.get('mrt'),'h')
                out=[]
                if q: out.append('q = '+q)
                if w: out.append('w = '+w)
                if t: out.append('t = '+t)
                return out or [mt]
            if mt == 'Probability':
                q=_fmt(e.get('prob')) or _fmt(e.get('alloc'))
                return ['q = '+q] if q else [mt]
            if mt == 'Multiplicity':
                m=_fmt(e.get('alloc'))
                return ['m = '+m] if m else [mt]
            if mt == 'True': return ['q = 1']
            if mt == 'False': return ['q = 0']
            q=_fmt(e.get('prob')) or _fmt(e.get('alloc'))
            return ['q = '+q] if q else [mt]
        if mt in ('ConstantRate','Fixed'):
            out=[]
            w=_fmt(e.get('alloc'),'/h'); s=_fmt(e.get('mrt'),'h')
            if w: out.append('w = '+w)
            if s: out.append('sdt = '+s)
            return out or ['w = ?']
        if mt=='Probability': return ['q = '+_fmt(e.get('alloc'))]
        if mt=='Multiplicity': return ['m = '+_fmt(e.get('alloc'))]
        if mt=='True': return ['q = 1']
        if mt=='False': return ['q = 0']
        return [mt]
    def _value_box(self,x,y,w,h,lines):
        self.rect(x,y,w,h)
        self.texts(lines, x+w/2, y+17, 'val')
    def _op_value_box_geometry(self,x,y,w,h):
        box_h = 58 if getattr(self,'pure',False) else 42
        box_w = _GATE_W
        box_x = x + w/2 - box_w/2
        box_y = y + h/2 - box_h/2
        return box_x, box_y, box_w, box_h
    def and_shape(self,x,y,w,h,lines):
        box_x, box_y, box_w, box_h = self._op_value_box_geometry(x,y,w,h)
        cx = x + w/2
        top = y + 8
        bottom = y + h - 8
        r = min(w*0.50, (bottom-top)*0.55)
        sx = cx - r
        d = (f'M {sx:.1f} {top+r:.1f} ' f'A {r:.1f} {r:.1f} 0 0 1 {sx+2*r:.1f} {top+r:.1f} ' f'L {sx+2*r:.1f} {bottom:.1f} ' f'L {sx:.1f} {bottom:.1f} Z')
        self.svg.append(f'<path d="{d}" fill="{_FILL}" stroke="{_STROKE}" stroke-width="1.5"/>')
        self._value_box(box_x, box_y, box_w, box_h, lines)
    def or_shape(self,x,y,w,h,lines):
        box_x, box_y, box_w, box_h = self._op_value_box_geometry(x,y,w,h)
        sx = x + 8
        sw = w - 16
        top = y + 6
        bottom = y + h - 6
        d = (f'M {sx:.1f} {bottom:.1f} ' f'C {sx+sw*.14:.1f} {top:.1f}, {sx+sw*.86:.1f} {top:.1f}, {sx+sw:.1f} {bottom:.1f} ' f'C {sx+sw*.68:.1f} {y+h*.61:.1f}, {sx+sw*.32:.1f} {y+h*.61:.1f}, {sx:.1f} {bottom:.1f} Z')
        self.svg.append(f'<path d="{d}" fill="{_FILL}" stroke="{_STROKE}" stroke-width="1.5"/>')
        self._value_box(box_x, box_y, box_w, box_h, lines)
    def draw_gate(self,gid,root,path=None):
        if path is None:
            path = ((0, gid),)
        x,y,w,h=self.pos[('gate',path)]; self.rect(x,y,w,_TITLE_H); self.texts(_wrap(self.gates.get(gid,{}).get('label',gid),30,4),x+w/2,y+24)
        self.rect(x,y+_TITLE_H+7,w,_ID_H); self.texts([gid],x+w/2,y+_TITLE_H+25,'small'); self.line(x+w/2,y+_TITLE_H,x+w/2,y+_TITLE_H+7)
        if gid not in self.children: return
        ox,oy,ow,oh=self.pos[('op',path)]
        is_or = self.gtypes.get(gid)=='OR'
        draw_h = oh*1.5 if is_or else oh
        op_y = oy - 42 if is_or else oy
        self.line(x+w/2,y+_TITLE_H+7+_ID_H,x+w/2,op_y+draw_h)
        (self.and_shape if self.gtypes.get(gid)=='AND' else self.or_shape)(ox,op_y,ow,draw_h,self.op_lines(gid))
        tops=[]
        for i,ch in enumerate(self.children.get(gid,[])):
            child_path = self._child_path(path, i, ch)
            if ch in self.grows and not (self.gates.get(ch,{}).get('is_page') and ch!=root):
                self.draw_gate(ch,root,child_path); cx,cy,cw,chh=self.pos[('gate',child_path)]; tops.append((cx+cw/2,cy))
            else:
                self.draw_leaf(ch,root,child_path); lx,ly,lw,lh=self.pos[('leaf',child_path)]; tops.append((lx+lw/2,ly))
        if tops:
            op_bottom = op_y + draw_h
            bus=min(t[1] for t in tops)-28; self.line(ox+ow/2,op_bottom,ox+ow/2,bus); self.line(min(t[0] for t in tops),bus,max(t[0] for t in tops),bus)
            for tx,ty in tops: self.line(tx,bus,tx,ty)
    def draw_leaf(self,nid,root,path):
        x,y,w,h=self.pos[('leaf',path)]
        if nid in self.grows:
            href=self.page_files.get(nid,'')
            tri_h = _OP_H
            pts=f'{x+w/2:.1f},{y:.1f} {x+w*.92:.1f},{y+tri_h:.1f} {x+w*.08:.1f},{y+tri_h:.1f}'
            if href: self.svg.append(f'<a href="{self.esc(href)}">')
            self.svg.append(f'<polygon points="{pts}" fill="{_FILL}" stroke="{_STROKE}" stroke-width="1.5"/>')
            box_w = _GATE_W
            id_h = _ID_H
            val_h = 58 if getattr(self,'pure',False) else 42
            gap = 16
            group_h = id_h + gap + val_h
            box_x = x + w/2 - box_w/2
            id_y = y + tri_h/2 - group_h/2
            val_y = id_y + id_h + gap
            self.rect(box_x, id_y, box_w, id_h)
            self.texts([nid], box_x + box_w/2, id_y + 18, 'small')
            self._value_box(box_x, val_y, box_w, val_h, self.transfer_value_lines(nid))
            if href: self.svg.append('</a>')
            return
        e=self.events.get(nid,{})
        self.rect(x,y,w,_TITLE_H); self.texts(_wrap(e.get('label',nid),30,4),x+w/2,y+24)
        self.rect(x,y+_TITLE_H+7,w,_ID_H); self.texts([nid],x+w/2,y+_TITLE_H+25,'small'); self.line(x+w/2,y+_TITLE_H,x+w/2,y+_TITLE_H+7)
        cx=x+w/2; cy=y+_TITLE_H+_ID_H+_EVENT_VALUE_H/2+26; r=_EVENT_VALUE_H/2
        self.svg.append(f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="{r:.1f}" fill="{_FILL}" stroke="{_STROKE}" stroke-width="1.5"/>')
        box_w = _GATE_W
        box_h = 58 if getattr(self,'pure',False) else 42
        box_x = cx - box_w/2
        box_y = cy - box_h/2
        self._value_box(box_x, box_y, box_w, box_h, self.event_lines(nid))
        self.line(cx,y+_TITLE_H+7+_ID_H,cx,cy-r)

def build_diagrams(xlsx, diag_prefix=None, output_dir=None, pure=False):
    src=Path(xlsx).expanduser().resolve(); out_dir=Path(output_dir).expanduser().resolve() if output_dir else Path(src.stem).resolve(); out_dir.mkdir(parents=True,exist_ok=True)
    dst=out_dir/src.name
    if src.resolve()!=dst.resolve(): _shutil.copy2(src,dst)
    data=_diagram_data(str(src), pure=pure); children,gtypes,erows,grows,events,gates,roots=data[:7]
    pages=[]
    for gid in roots:
        if gid not in pages: pages.append(gid)
    for gid in grows:
        if gates.get(gid,{}).get('is_page') and gid in children and gid not in pages: pages.append(gid)
    prefix=diag_prefix or src.stem; page_files={gid:f"{prefix}_fta_{'main' if i==0 else gid}.svg" for i,gid in enumerate(pages)}
    r=FtaSvg(data,page_files)
    for gid in pages: r.render(gid,out_dir/page_files[gid])
    links=''.join(f'<li><a href="{_html.escape(page_files[g])}">{_html.escape(g)}</a></li>' for g in pages)
    sections=''.join(f'<h2>{_html.escape(g)}</h2><object type="image/svg+xml" data="{_html.escape(page_files[g])}" style="width:100%;min-height:850px;border:1px solid #ddd"></object>' for g in pages)
    html_file=out_dir/f'{prefix}_fta.html'; html_file.write_text(f"<!doctype html><html><head><meta charset='utf-8'><title>FTA diagrams</title></head><body><h1>FTA diagrams</h1><p>Workbook: <code>{_html.escape(src.name)}</code></p><ul>{links}</ul>{sections}</body></html>",encoding='utf-8')
    return str(html_file.resolve())
# --- End native SVG FTA renderer v7 -----------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description=(
            "Allocative FTA Excel updater. By default the allocative model is Markov. "
            "Use --pure to generate the pure-probability allocative workbook. "
            "Use --diag only when diagrams are required. "
            "When --diag is used without --markov/--pure, diagrams are generated for both models."
        )
    )
    ap.add_argument("excel")

    # Backward-compatible / explicit model options.
    # --alloc is kept for older command lines, but allocative mode is now the default.
    ap.add_argument("--alloc", action="store_true", help="Backward-compatible no-op: allocative mode is now the default.")
    ap.add_argument("--markov", action="store_true", help="Use Markov allocative formulas. This is the default when --pure is not given.")
    ap.add_argument("--pure", action="store_true", help="Use pure-probability allocative formulas instead of Markov formulas.")

    # Output / diagram options.
    ap.add_argument("--diag", action="store_true", help="Generate FTA diagrams in addition to the Excel workbook.")
    ap.add_argument("--alloc-out", default=None, help="Output path for the Markov workbook. Kept for backward compatibility.")
    ap.add_argument("--markov-out", default=None, help="Alias for --alloc-out.")
    ap.add_argument("--pure-out", default=None, help="Output path for the pure-probability workbook.")
    ap.add_argument("--diag-prefix", default=None)
    ap.add_argument("--diag-dir", default=None)
    args = ap.parse_args()

    if args.pure and args.markov:
        raise SystemExit("Choose either --pure or --markov, not both.")

    # Model selection rules:
    #   - --pure or --markov explicitly scopes work to one model.
    #   - without explicit model flags, Markov remains the default workbook mode.
    #   - with --diag and no explicit model flag, generate diagrams for both models.
    explicit_model = "pure" if args.pure else ("markov" if args.markov else None)

    input_path = Path(args.excel).expanduser().resolve()
    out_dir = args.diag_dir
    if out_dir is None:
        out_dir = str(input_path.with_suffix(""))
    Path(out_dir).mkdir(parents=True, exist_ok=True)

    def build_markov() -> str:
        markov_out = args.markov_out or args.alloc_out
        if markov_out is None:
            markov_out = str(Path(out_dir) / (input_path.stem + "_alloc" + input_path.suffix))
        path = build_markov_workbook(str(input_path), markov_out)
        print("OK markov: " + path)
        return path

    def build_pure() -> str:
        pure_out = args.pure_out
        if pure_out is None:
            pure_out = str(Path(out_dir) / (input_path.stem + "_pure" + input_path.suffix))
        path = build_pure_workbook(str(input_path), pure_out)
        print("OK pure: " + path)
        return path

    outputs: Dict[str, str] = {}

    if explicit_model == "pure":
        outputs["pure"] = build_pure()
    elif explicit_model == "markov":
        outputs["markov"] = build_markov()
    elif args.diag:
        # For diagram runs without explicit model, generate all model outputs.
        outputs["markov"] = build_markov()
        outputs["pure"] = build_pure()
    else:
        # Non-diagram default remains Markov for backward compatibility.
        outputs["markov"] = build_markov()

    # Diagrams are generated only when explicitly requested.
    if args.diag:
        if explicit_model == "pure":
            diag_targets = [("pure", outputs["pure"])]
        elif explicit_model == "markov":
            diag_targets = [("markov", outputs["markov"])]
        else:
            diag_targets = [("markov", outputs["markov"]), ("pure", outputs["pure"])]

        for model_name, workbook_path in diag_targets:
            # Avoid filename collisions when generating both model diagram sets.
            if args.diag_prefix:
                prefix = args.diag_prefix if len(diag_targets) == 1 else f"{args.diag_prefix}_{model_name}"
            else:
                prefix = None
            html_out = build_diagrams(
                workbook_path,
                prefix,
                output_dir=out_dir,
                pure=(model_name == "pure"),
            )
            print(f"OK diag ({model_name}): {html_out}")

if __name__ == "__main__": main()

