import calendar
import os
import re
import secrets
from typing import Any

from fastapi import APIRouter, Depends, Header, HTTPException, Query
from fastapi.responses import Response
from sqlalchemy.orm import Session

from app.core.database import get_db
from app.core.security import decode_access_token
from app.db.models import User
from app.services.monthly_juja_cluster_report_builder import build_monthly_juja_cluster_excel_report
from app.services.monthly_mobile_report_builder import build_monthly_mobile_excel_report


router = APIRouter()


def require_admin_password(
    x_admin_password: str | None,
    authorization: str | None = None,
) -> None:
    if authorization and authorization.startswith("Bearer "):
        token = authorization.split(" ")[1]
        payload = decode_access_token(token)
        if payload:
            role = payload.get("role")
            username = payload.get("sub")
            if role in ["admin", "developer", "viewer", "user", "duty_manager"] or username == "admin":
                return

    configured_password = os.getenv("ADMIN_PASSWORD")
    if not configured_password:
        return

    if not x_admin_password or not secrets.compare_digest(
        x_admin_password,
        configured_password,
    ):
        raise HTTPException(
            status_code=401,
            detail="Invalid admin password or session token.",
        )


def get_authenticated_user(
    authorization: str | None,
    db: Session,
) -> User | None:
    if authorization and authorization.startswith("Bearer "):
        token = authorization.split(" ")[1]
        payload = decode_access_token(token)
        if payload:
            username = payload.get("sub")
            if username:
                return db.query(User).filter(User.username == username).first()
    return None


def resolve_user_station(user: User | None) -> str | None:
    if not user:
        return None
    raw = (user.station or user.username or "").strip().lower()
    if "kanyonyo" in raw:
        return "KANYONYO"
    if "isinya" in raw:
        return "ISINYA"
    if "athi" in raw:
        return "ATHI RIVER"
    if "gilgil" in raw:
        return "GILGIL"
    if "suswa" in raw:
        return "SUSWA"
    if "juja" in raw:
        return "JUJA"
    if user.station:
        return user.station.strip().upper()
    return None


@router.get("/reports/monthly/generate")
async def generate_monthly_report(
    station: str = Query(..., description="Mobile Station name e.g. Juja mobile"),
    year: int = Query(..., ge=2020, le=2050),
    month: int = Query(..., ge=1, le=12),
    report_type: str = Query("mobile_summary", description="Type of monthly report to generate"),
    format: str = Query("excel", pattern="^(excel)$"),
    x_admin_password: str | None = Header(default=None),
    authorization: str | None = Header(default=None),
    db: Session = Depends(get_db),
):
    require_admin_password(x_admin_password, authorization)

    auth_user = get_authenticated_user(authorization, db)
    if auth_user and auth_user.role != "admin":
        assigned_station = resolve_user_station(auth_user)
        if assigned_station:
            req_station = "juja" if report_type == "juja_cluster_invoicing" else station
            req_clean = re.sub(r"[^a-z0-9]", "", req_station.lower()).replace("mobile", "")
            assigned_clean = re.sub(r"[^a-z0-9]", "", assigned_station.lower()).replace("mobile", "")
            if req_clean != assigned_clean and assigned_clean not in req_clean:
                raise HTTPException(
                    status_code=403,
                    detail=f"Access denied: You are only authorized to generate monthly reports for {assigned_station} station."
                )

    month_name = calendar.month_name[month]
    station_clean = station.strip().replace(" ", "_").title()

    if report_type == "mobile_summary":
        buffer = build_monthly_mobile_excel_report(station, year, month)
        filename = f"{station_clean}_Monthly_Summary_{month_name}_{year}.xlsx"
        return Response(
            content=buffer.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    elif report_type == "juja_cluster_invoicing":
        buffer = build_monthly_juja_cluster_excel_report(year, month)
        filename = f"Monthly JUJA Cluster Invoicing Report {month_name.upper()} {year}.xlsx"
        return Response(
            content=buffer.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    else:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported report type: {report_type}. Supported types: 'mobile_summary', 'juja_cluster_invoicing'."
        )
