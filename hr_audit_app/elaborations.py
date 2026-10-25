# -*- coding: utf-8 -*-
"""Cau phan tich RBA bo sung cho finding ngan (song ngu EN/VI thuan)."""
try:
    from elaborations_full import ELABORATIONS_FULL  # noqa
except ImportError:
    try:
        from .elaborations_full import ELABORATIONS_FULL  # noqa
    except ImportError:
        ELABORATIONS_FULL = {}
_ELABORATE_MIN_LEN = 150


def _elaborate_key(audit_id, get_item):
    item = get_item(audit_id)
    if not item:
        return None
    aid = (audit_id or "").upper().strip()
    if aid == "AL504":
        return "disability"
    if aid == "AL503":
        return "religion"
    if aid == "AL502":
        return "discipline"
    sec = (item.get("rba_section") or "").strip()
    if sec.startswith("A.M"):
        return "labor_mgmt"
    if sec.startswith("A1"):
        return "forced_labor"
    if sec.startswith("A2"):
        return "young_workers"
    if sec.startswith("A3"):
        return "working_hours"
    if sec.startswith("A4"):
        return "wages_benefits"
    if sec.startswith("A5"):
        return "non_discrimination"
    if sec.startswith("A6"):
        return "freedom_association"
    if sec.startswith("D6"):
        return "privacy"
    if sec.startswith("D"):
        return "ethics_mgmt"
    if sec.startswith("E"):
        return "supply_chain"
    return "general"
