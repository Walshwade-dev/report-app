from fastapi import APIRouter, HTTPException, Depends, Query, Header, Body
from typing import Optional, List, Dict, Any
from datetime import datetime, date
import calendar
import uuid

router = APIRouter(prefix="/api/mobile-checklist", tags=["Mobile Checklist"])

# In-memory storage with initial baseline seed data
CHECKLIST_ENTRIES: List[Dict[str, Any]] = [
    {
        "id": "chk-03092026-day",
        "date": "2026-09-03",
        "shift": "Day Shift Mobile Operations",
        "station": "Mobile 1 Weighbridge",
        "scales_sno": "6451",
        "technician_name": "John Kimani (Technician)",
        "approved_by": "Peter Njoroge (Duty Manager)",
        "status": "Approved",
        "approval_date": "2026-09-03T18:00:00",
        "mobile_scale_gvw": 12440,
        "multideck_scale_gvw": 11960,
        "variance_gvw": 480,
        "variance_comment": "ALERT: Scale weight variance of 480kg exceeds 200kg threshold. Recalibration and technical inspection recommended.",
        "technical_checks": {
          "plate_screws_tight": True,
          "surface_clean": True,
          "battery_charge": 95,
          "dry_operation": True,
          "cables_intact": True
        },
        "vehicle_reg": "KCM 494U / KCF 951Q",
        "doc_ref": "DNK/AFRKE/LV3/WBS/018/FM03",
        "created_at": "2026-09-03T07:30:00"
    }
]

EQUIPMENT_ISSUES: List[Dict[str, Any]] = [
    {
        "id": "iss-001",
        "equipment_name": "10m load scale connecting Cable (E Cable 6920.3)",
        "issue_description": "Intermittent signal loss on channel 2 due to connector pin wear.",
        "reported_date": "2026-09-02",
        "reported_by": "Duty Manager",
        "status": "In Progress",
        "severity": "Medium",
        "assigned_handler": "Alex Technical Officer",
        "resolution_notes": "Replacement cable ordered from warehouse. Secondary backup cable deployed.",
        "created_at": "2026-09-02T10:15:00"
    },
    {
        "id": "iss-002",
        "equipment_name": "12Vdc to 240Vac Power Inverter",
        "issue_description": "Fuse blown during heavy night shift charging.",
        "reported_date": "2026-08-30",
        "reported_by": "Night Operator",
        "status": "Resolved",
        "severity": "High",
        "assigned_handler": "Electrical Maintenance Team",
        "resolution_notes": "15A fuse replaced and inverter bench-tested successfully.",
        "created_at": "2026-08-30T22:40:00"
    }
]

DELIVERY_NOTES: List[Dict[str, Any]] = [
    {
        "id": "del-001",
        "delivery_date": "2026-08-25",
        "item_name": "WL 108 Wheel Load Scale 15-Ton (Set of 2)",
        "serial_numbers": "6451-A, 6451-B",
        "condition_state": "Brand New - Factory Calibrated",
        "intended_purpose": "Mobile Operations Primary Weighing Unit",
        "delivered_by": "Avery Weights Kenya Ltd",
        "received_by": "Technical Manager",
        "pdf_url": "/downloads/delivery_note_wl108_20260825.pdf",
        "created_at": "2026-08-25T14:20:00"
    }
]


def cleanup_expired_entries():
    """Purge checklist entries older than the current month's total days retention window."""
    today = date.today()
    days_in_month = calendar.monthrange(today.year, today.month)[1]
    
    global CHECKLIST_ENTRIES
    valid_entries = []
    for entry in CHECKLIST_ENTRIES:
        try:
            entry_date = datetime.strptime(entry["date"], "%Y-%m-%d").date()
            delta_days = (today - entry_date).days
            if delta_days <= days_in_month:
                valid_entries.append(entry)
        except Exception:
            valid_entries.append(entry)
    CHECKLIST_ENTRIES = valid_entries


@router.get("/entries")
def get_checklist_entries():
    cleanup_expired_entries()
    return {"entries": CHECKLIST_ENTRIES, "total": len(CHECKLIST_ENTRIES)}


@router.post("/entries")
def create_checklist_entry(payload: Dict[str, Any] = Body(...)):
    mobile_gvw = float(payload.get("mobile_scale_gvw", 0))
    multideck_gvw = float(payload.get("multideck_scale_gvw", 0))
    variance = abs(mobile_gvw - multideck_gvw)
    
    if variance <= 200:
        variance_comment = f"Variance of {variance:.0f}kg is within acceptable calibration tolerance (<= 200kg)."
    else:
        variance_comment = f"ALERT: Scale weight variance of {variance:.0f}kg exceeds 200kg threshold. Recalibration and technical inspection recommended."

    new_entry = {
        "id": f"chk-{uuid.uuid4().hex[:8]}",
        "date": payload.get("date", str(date.today())),
        "shift": payload.get("shift", "Day Shift Mobile Operations"),
        "station": payload.get("station", "Mobile 1 Weighbridge"),
        "scales_sno": payload.get("scales_sno", "6451"),
        "technician_name": payload.get("technician_name", "Technician Officer"),
        "approved_by": None,
        "status": "Submitted",
        "mobile_scale_gvw": mobile_gvw,
        "multideck_scale_gvw": multideck_gvw,
        "variance_gvw": variance,
        "variance_comment": variance_comment,
        "items": payload.get("items", []),
        "technical_checks": payload.get("technical_checks", {}),
        "vehicle_reg": payload.get("vehicle_reg", ""),
        "doc_ref": "DNK/AFRKE/LV3/WBS/018/FM03",
        "created_at": datetime.utcnow().isoformat()
    }
    
    CHECKLIST_ENTRIES.insert(0, new_entry)
    cleanup_expired_entries()
    return {"message": "Checklist saved successfully", "entry": new_entry}


@router.patch("/entries/{entry_id}/approve")
def approve_checklist_entry(entry_id: str, payload: Dict[str, Any] = Body(...)):
    for entry in CHECKLIST_ENTRIES:
        if entry["id"] == entry_id:
            entry["status"] = "Approved"
            entry["approved_by"] = payload.get("approved_by", "Duty Manager")
            entry["approval_date"] = datetime.utcnow().isoformat()
            return {"message": "Checklist approved and archived", "entry": entry}
            
    raise HTTPException(status_code=404, detail="Checklist entry not found")


@router.get("/issues")
def get_equipment_issues():
    return {"issues": EQUIPMENT_ISSUES, "total": len(EQUIPMENT_ISSUES)}


@router.post("/issues")
def create_equipment_issue(payload: Dict[str, Any] = Body(...)):
    new_issue = {
        "id": f"iss-{uuid.uuid4().hex[:6]}",
        "equipment_name": payload.get("equipment_name", "Mobile Equipment"),
        "issue_description": payload.get("issue_description", ""),
        "reported_date": payload.get("reported_date", str(date.today())),
        "reported_by": payload.get("reported_by", "Duty Manager"),
        "status": payload.get("status", "Pending"),
        "severity": payload.get("severity", "Medium"),
        "assigned_handler": payload.get("assigned_handler", "Unassigned"),
        "resolution_notes": payload.get("resolution_notes", ""),
        "created_at": datetime.utcnow().isoformat()
    }
    EQUIPMENT_ISSUES.insert(0, new_issue)
    return {"message": "Equipment issue logged", "issue": new_issue}


@router.patch("/issues/{issue_id}")
def update_equipment_issue(issue_id: str, payload: Dict[str, Any] = Body(...)):
    for issue in EQUIPMENT_ISSUES:
        if issue["id"] == issue_id:
            if "status" in payload:
                issue["status"] = payload["status"]
            if "assigned_handler" in payload:
                issue["assigned_handler"] = payload["assigned_handler"]
            if "resolution_notes" in payload:
                issue["resolution_notes"] = payload["resolution_notes"]
            return {"message": "Equipment issue updated", "issue": issue}
            
    raise HTTPException(status_code=404, detail="Equipment issue not found")


@router.get("/delivery-notes")
def get_delivery_notes():
    return {"notes": DELIVERY_NOTES, "total": len(DELIVERY_NOTES)}


@router.post("/delivery-notes")
def create_delivery_note(payload: Dict[str, Any] = Body(...)):
    new_note = {
        "id": f"del-{uuid.uuid4().hex[:6]}",
        "delivery_date": payload.get("delivery_date", str(date.today())),
        "item_name": payload.get("item_name", ""),
        "serial_numbers": payload.get("serial_numbers", "N/A"),
        "condition_state": payload.get("condition_state", "Good"),
        "intended_purpose": payload.get("intended_purpose", ""),
        "delivered_by": payload.get("delivered_by", ""),
        "received_by": payload.get("received_by", ""),
        "pdf_url": payload.get("pdf_url", ""),
        "created_at": datetime.utcnow().isoformat()
    }
    DELIVERY_NOTES.insert(0, new_note)
    return {"message": "Delivery note saved", "note": new_note}
