from __future__ import annotations

import json
import sys
from pathlib import Path

# Core engine is kept at package root so it can still be used independently from CLI.
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import monitoring_automation_v1 as core

# Additional fields needed by dashboards; adding them does not change configured finding logic.
EXTRA_FIELDS = {
    "assistance_mode": ["What types of WFP assistance has your household received?"],
    "hh_food": ["HHAsstMode/Food"],
    "hh_cash": ["HHAsstMode/Cash"],
    "interviewee_sex": ["1.2 Sex of the interviewee"],
    "is_hh_head": ["1.3 Are you the head of the household?"],
    "hh_head_sex": ["1.4 What is the Sex of the household head?"],
    "cfm_awareness_general": ["If you wanted to ask a question, get more information, make a complaint/appeal, report misconduct or provide feedback, do you know what to do/who to contact?"],
    "cfm_channels_general": ["What feedback mechanisms are available?"],
    "cfm_helpdesk_access": ["Does the FDP have a functional and accessible help desk?"],
    "tsfp_cfm_awareness": ["1. Do you know how to present your complaint or feedback regarding the TSFP in case of need?"],
    "tsfp_cfm_channel": ["1.2 If yes, to whom you would primarily present your complaint?"],
    "tsfp_cfm_usage": ["1.3 If yes, have you ever utilized the available feedback mechanisms?"],
    "res_cfm_awareness": ["K3. Do you know where to report if you have a complaint?"],
    "res_cfm_channel": ["K4. If yes, where do you report?"],
    "res_cfm_usage": ["K5. Have you ever reported a complaint using the WFP/CP hotline before?"],
    "res_cfm_resolution_time": ["K7. How long did it take to resolve your issue?"],
}
for k, v in EXTRA_FIELDS.items():
    core.BASE_FIELDS.setdefault(k, v)


def run_analysis(moda_paths: list[str], reporting_month: str, sub_office: str = "Jijiga") -> dict:
    rows, schema, intake = core.load_inputs(moda_paths, reporting_month, sub_office)
    indicators = core.indicator_summary(rows)
    detail = core.detailed_findings(rows, reporting_month)
    agg = core.aggregate_actions(detail, indicators, reporting_month)
    dq = core.data_quality(rows, schema, intake, reporting_month)
    return {"rows": rows, "schema": schema, "intake": intake, "indicators": indicators, "detail": detail, "agg": agg, "dq": dq}
