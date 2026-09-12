from __future__ import annotations

import logging
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from app.services.report_session_store import ReportSession, report_session_store

logger = logging.getLogger(__name__)

STATION_NAMES = {
    "Juja": "Juja Weighbridge",
    "Kanyonyo": "Kanyonyo",
    "Athi River": "Athi River",
    "Gilgil": "Gilgil",
    "Isinya": "Isinya",
    "Suswa": "Suswa",
}


def _classify_station(station_name: str | None) -> str | None:
    if not station_name:
        return None
    name = station_name.lower()
    if "juja" in name:
        return "Juja"
    if "kanyonyo" in name:
        return "Kanyonyo"
    if "athi" in name:
        return "Athi River"
    if "gilgil" in name:
        return "Gilgil"
    if "isinya" in name:
        return "Isinya"
    if "suswa" in name:
        return "Suswa"
    return None


def _is_bound_a(station_code: str | None, bound_name: str | None) -> bool:
    if not bound_name:
        return True
    bound = bound_name.lower()
    if not station_code:
        return any(
            k in bound
            for k in ["bound a", "incoming", "thika", "mombasa", "mwingi", "kajiado", "narok"]
        )

    st = station_code.lower()
    if "juja" in st:
        return "thika" in bound or "bound a" in bound or "incoming" in bound
    elif "athi" in st:
        return "mombasa" in bound or "bound a" in bound or "incoming" in bound
    elif "gilgil" in st:
        return "nairobi" in bound or "bound a" in bound or "incoming" in bound
    elif "kanyonyo" in st:
        return True
    elif "isinya" in st:
        return "kajiado" in bound or "bound a" in bound or "incoming" in bound
    elif "suswa" in st:
        return "narok" in bound or "bound a" in bound or "incoming" in bound

    return any(
        k in bound
        for k in ["bound a", "incoming", "thika", "mombasa", "mwingi", "kajiado", "narok"]
    )


def _daily_hour_total(session: ReportSession, column: str) -> int | None:
    if "daily_hour" not in session.dataframes or session.sections.get("daily_hour", {}).get("status") != "ready":
        sec = session.sections.get("daily_hour", {})
        if isinstance(sec, dict) and "summary" in sec:
            summary = sec["summary"]
            if isinstance(summary, dict) and column in summary:
                try:
                    return int(summary[column])
                except Exception:
                    pass
        return None

    daily_df = session.dataframes["daily_hour"]
    if "DATE" not in daily_df.columns or column not in daily_df.columns:
        return None

    totals_mask = daily_df["DATE"].astype(str).str.strip().str.lower().eq("totals")
    if not totals_mask.any():
        return None

    try:
        return int(daily_df.loc[totals_mask].iloc[-1].get(column, 0))
    except Exception:
        return None


def _clean_cargo_name(cargo_raw: Any) -> str:
    if not cargo_raw:
        return "UNKNOWN"
    val = str(cargo_raw).strip().upper()
    if not val or val in ["NAN", "NONE", "N/A", "-", "NULL", "0"]:
        return "UNKNOWN"
    # Normalize common synonyms
    if "CEMENT" in val and "50" in val:
        return "CEMENT (50KG BAGS)"
    elif "CEMENT" in val:
        return "CEMENT"
    elif "BUILDING" in val or "B MATERIALS" in val or "CONSTRUCTION" in val:
        return "BUILDING & CONSTRUCTION MATERIALS"
    elif "PASSENGER" in val or "PSV" in val:
        return "PSV / PASSENGERS"
    elif "OVACADO" in val or "AVOCADO" in val:
        return "AVOCADO / HORTICULTURE"
    elif "FARM" in val or "PRODUCE" in val:
        return "FARM PRODUCE"
    elif "ASSORTED" in val:
        return "ASSORTED GOODS"
    return val


def _read_processed_pickle(storage_root: Path, report_id: str, pkl_name: str) -> pd.DataFrame | None:
    path = storage_root / "processed" / report_id / f"{pkl_name}.pkl"
    if path.exists():
        try:
            df = pd.read_pickle(path)
            if isinstance(df, pd.DataFrame) and not df.empty:
                return df
        except Exception:
            logger.debug("Failed to read pickle %s for %s", pkl_name, report_id)
    return None


def get_analytics_details(
    station: str | None = None,
    month: str | None = None,
    year: int | None = None,
) -> dict[str, Any]:
    """
    Computes comprehensive analytics scoped strictly to the selected month and year
    (from the first day to the end of that month).
    Excludes all DMS names and team names.
    """
    sessions_with_mtime = report_session_store.list_all_sessions()
    all_sessions = [s for s, _ in sessions_with_mtime]
    storage_root = report_session_store.storage_root

    # 1. Discover all available months (sorted newest first)
    all_months_set: set[str] = set()
    for s in all_sessions:
        if s.report_date and len(s.report_date) >= 7:
            all_months_set.add(s.report_date[:7])

    available_months = sorted(all_months_set, reverse=True)

    # 2. Determine target month string "YYYY-MM"
    target_month: str | None = None

    if month and len(month.strip()) == 7 and "-" in month:
        target_month = month.strip()
    elif month and year:
        try:
            m_int = int(month.strip())
            target_month = f"{year:04d}-{m_int:02d}"
        except Exception:
            target_month = None
    elif month and not year:
        try:
            m_int = int(month.strip())
            current_y = datetime.now().year
            target_month = f"{current_y:04d}-{m_int:02d}"
        except Exception:
            target_month = None

    if not target_month:
        now = datetime.now()
        current_cal_month = now.strftime("%Y-%m")
        if current_cal_month in available_months:
            target_month = current_cal_month
        elif available_months:
            target_month = available_months[0]
        else:
            target_month = current_cal_month

    # Format human-readable target month label (e.g. "September 2026")
    try:
        dt_obj = datetime.strptime(target_month, "%Y-%m")
        selected_month_label = dt_obj.strftime("%B %Y")
        selected_month_name = dt_obj.strftime("%B")
        selected_year = dt_obj.year
        selected_month_num = dt_obj.month
    except Exception:
        selected_month_label = target_month
        selected_month_name = target_month
        selected_year = datetime.now().year
        selected_month_num = datetime.now().month

    # Format list of available month objects for the frontend selector
    available_months_formatted = []
    for ym in available_months:
        try:
            d = datetime.strptime(ym, "%Y-%m")
            available_months_formatted.append({
                "value": ym,
                "label": d.strftime("%B %Y"),
                "year": d.year,
                "month": d.month,
            })
        except Exception:
            available_months_formatted.append({
                "value": ym,
                "label": ym,
                "year": selected_year,
                "month": selected_month_num,
            })

    # 3. Filter sessions strictly within target month (YYYY-MM-01 to end of month)
    month_sessions = [s for s in all_sessions if s.report_date and s.report_date.startswith(target_month)]

    # 4. Scope station
    target_code = _classify_station(station) or (station.strip() if station else "Juja")
    target_lower = target_code.lower()

    bound_a_name = "Bound A"
    bound_b_name = "Bound B"
    if "juja" in target_lower:
        bound_a_name = "Thika Bound"
        bound_b_name = "Nairobi Bound"
    elif "athi" in target_lower:
        bound_a_name = "Mombasa Bound"
        bound_b_name = "Nairobi Bound"
    elif "gilgil" in target_lower:
        bound_a_name = "Nairobi Bound"
        bound_b_name = "Nakuru Bound"
    elif "kanyonyo" in target_lower:
        bound_a_name = "Nairobi Bound"
        bound_b_name = ""
    elif "isinya" in target_lower:
        bound_a_name = "Kajiado Bound"
        bound_b_name = "Nairobi Bound"
    elif "suswa" in target_lower:
        bound_a_name = "Narok Bound"
        bound_b_name = "Nairobi Bound"

    station_month_sessions = [
        s for s in month_sessions
        if (_classify_station(s.station or s.weighbridge_name) or (s.station or s.weighbridge_name or "").strip()).lower() == target_lower
    ]

    # 5. Static KPIs (for the month)
    station_bound_a_traffic = 0
    station_bound_b_traffic = 0
    station_bound_a_cases = 0
    station_bound_b_cases = 0
    station_total_called = 0
    station_total_compliant = 0
    station_overloads_intercepted = 0

    daily_traffic: dict[str, dict[str, int]] = {}
    daily_cases: dict[str, dict[str, int]] = {}

    for s in station_month_sessions:
        is_a = _is_bound_a(target_code, s.bound)

        weighed = _daily_hour_total(s, "X") or 0
        if is_a:
            station_bound_a_traffic += weighed
        else:
            station_bound_b_traffic += weighed

        cases = s.manual_inputs.get("cases_cleared_in_court", 0) or 0
        try:
            cases = int(cases)
        except Exception:
            cases = 0

        if is_a:
            station_bound_a_cases += cases
        else:
            station_bound_b_cases += cases

        called = _daily_hour_total(s, "C") or 0
        y = _daily_hour_total(s, "Y") or 0
        g = _daily_hour_total(s, "G") or 0
        overload_no_permit = max(y - g, 0)
        compliant = max(called - overload_no_permit, 0)

        station_total_called += called
        station_total_compliant += compliant
        station_overloads_intercepted += overload_no_permit

        try:
            day_str = s.report_date.split("-")[2]
        except Exception:
            continue

        if day_str not in daily_traffic:
            daily_traffic[day_str] = {"boundA": 0, "boundB": 0}
        if day_str not in daily_cases:
            daily_cases[day_str] = {"boundA": 0, "boundB": 0}

        if is_a:
            daily_traffic[day_str]["boundA"] += weighed
            daily_cases[day_str]["boundA"] += cases
        else:
            daily_traffic[day_str]["boundB"] += weighed
            daily_cases[day_str]["boundB"] += cases

    station_total_traffic = station_bound_a_traffic + station_bound_b_traffic
    station_total_cases = station_bound_a_cases + station_bound_b_cases
    station_compliance_rate = (station_total_compliant / station_total_called * 100) if station_total_called > 0 else 0.0

    traffic_data = []
    for day in sorted(daily_traffic.keys()):
        traffic_data.append({
            "day": day,
            "thikaBound": daily_traffic[day]["boundA"],
            "nairobiBound": daily_traffic[day]["boundB"],
            "boundA": daily_traffic[day]["boundA"],
            "boundB": daily_traffic[day]["boundB"],
        })

    court_cases_data = []
    for day in sorted(daily_cases.keys()):
        court_cases_data.append({
            "day": day,
            "thikaBound": daily_cases[day]["boundA"],
            "nairobiBound": daily_cases[day]["boundB"],
            "boundA": daily_cases[day]["boundA"],
            "boundB": daily_cases[day]["boundB"],
        })

    cross_station = {code: 0 for code in STATION_NAMES}
    for s in month_sessions:
        code = _classify_station(s.station or s.weighbridge_name)
        if code in cross_station:
            c = s.manual_inputs.get("cases_cleared_in_court", 0) or 0
            try:
                cross_station[code] += int(c)
            except Exception:
                pass

    cross_station_data = []
    for code, name in STATION_NAMES.items():
        cross_station_data.append({
            "name": name,
            "cases": cross_station[code],
            "active": code.lower() == target_lower,
        })

    # 6. Mobile Data Analytics: Routes Prone to Charging
    # Strictly exclude any DMS names or team names
    route_aggregates: dict[str, dict[str, Any]] = defaultdict(
        lambda: {
            "route": "",
            "routeType": "Patrol Route",
            "totalWeighed": 0,
            "chargedCount": 0,
            "warnedCount": 0,
            "legalCount": 0,
            "cargosCharged": Counter(),
            "datesActive": set(),
        }
    )

    total_mobile_weighed = 0
    total_mobile_charged = 0
    total_mobile_warned = 0
    total_mobile_legal = 0

    mobile_records_by_date_plate: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    cargo_overload_counts: Counter[str] = Counter()
    cargo_overload_kg_total: dict[str, float] = defaultdict(float)
    cargo_max_overload: dict[str, float] = defaultdict(float)

    for s in month_sessions:
        is_mobile_session = (
            s.sections.get("mobile_report", {}).get("status") == "ready"
            or bool(s.manual_inputs.get("mobile_report"))
            or "mobile" in (s.station or "").lower()
            or "mobile" in (s.bound or "").lower()
        )
        if not is_mobile_session:
            continue

        m_inputs = s.manual_inputs.get("mobile_report") or {}
        patrol_route_name = str(m_inputs.get("route") or "").strip().upper()
        if not patrol_route_name or patrol_route_name in ["UNKNOWN", "NONE", "N/A", "-"]:
            patrol_route_name = f"Mobile Patrol ({s.bound or 'Corridor'})"

        df = s.dataframes.get("mobile_report")
        if df is None:
            df = _read_processed_pickle(storage_root, s.report_id, "mobile_report")

        if df is not None and not df.empty:
            for _, row in df.iterrows():
                total_mobile_weighed += 1
                remarks = str(row.get("remarks") or "").strip().upper()
                diff_kg = float(row.get("gvw_difference_kg") or 0)
                cargo_clean = _clean_cargo_name(row.get("cargo"))
                plate = str(row.get("registration") or "").strip().upper()
                origin = str(row.get("origin") or "").strip().upper()
                dest = str(row.get("destination") or "").strip().upper()

                dt_val = str(row.get("date_time") or "")
                rec_date = dt_val.split(" ")[0] if dt_val and "-" in dt_val else s.report_date

                is_charged = "CHARG" in remarks or diff_kg > 2000
                is_warned = "WARN" in remarks or (0 < diff_kg <= 2000)

                if is_charged:
                    total_mobile_charged += 1
                elif is_warned:
                    total_mobile_warned += 1
                else:
                    total_mobile_legal += 1

                # Operational patrol route
                r_entry = route_aggregates[patrol_route_name]
                r_entry["route"] = patrol_route_name
                r_entry["routeType"] = "Patrol Operation"
                r_entry["totalWeighed"] += 1
                r_entry["datesActive"].add(rec_date)
                if is_charged:
                    r_entry["chargedCount"] += 1
                    if cargo_clean != "UNKNOWN":
                        r_entry["cargosCharged"][cargo_clean] += 1
                elif is_warned:
                    r_entry["warnedCount"] += 1
                else:
                    r_entry["legalCount"] += 1

                # Specific vehicle corridor
                if origin and dest and origin not in ["UNKNOWN", "NAN"] and dest not in ["UNKNOWN", "NAN"]:
                    corridor_name = f"{origin} ➔ {dest}"
                    c_entry = route_aggregates[corridor_name]
                    c_entry["route"] = corridor_name
                    c_entry["routeType"] = "Transport Corridor"
                    c_entry["totalWeighed"] += 1
                    c_entry["datesActive"].add(rec_date)
                    if is_charged:
                        c_entry["chargedCount"] += 1
                        if cargo_clean != "UNKNOWN":
                            c_entry["cargosCharged"][cargo_clean] += 1
                    elif is_warned:
                        c_entry["warnedCount"] += 1
                    else:
                        c_entry["legalCount"] += 1

                if is_charged and cargo_clean != "UNKNOWN":
                    cargo_overload_counts[cargo_clean] += 1
                    cargo_overload_kg_total[cargo_clean] += max(diff_kg, 0)
                    if diff_kg > cargo_max_overload[cargo_clean]:
                        cargo_max_overload[cargo_clean] = diff_kg

                if plate and plate not in ["NAN", "UNKNOWN", "NONE", "-"]:
                    total_gvw = float(row.get("total_gvw_kg") or row.get("gvw_kg") or 0)
                    mobile_records_by_date_plate[(rec_date, plate)].append({
                        "session_id": s.report_id,
                        "date": rec_date,
                        "time": dt_val.split(" ")[1] if " " in dt_val else "",
                        "registration": plate,
                        "station": s.station or "Mobile Weighbridge",
                        "bound": s.bound or "Mobile",
                        "cargo": cargo_clean,
                        "totalGvwKg": total_gvw,
                        "diffKg": diff_kg,
                        "remarks": remarks or ("CHARGED" if is_charged else "WARNED" if is_warned else "LEGAL"),
                        "origin": origin,
                        "destination": dest,
                    })

    routes_prone_to_charging = []
    for r_name, r_data in route_aggregates.items():
        tot = r_data["totalWeighed"]
        chg = r_data["chargedCount"]
        rate = round((chg / tot * 100), 1) if tot > 0 else 0.0
        top_cargos = [c for c, _ in r_data["cargosCharged"].most_common(3)]

        if chg >= 3 or rate >= 20.0:
            risk = "High Risk"
        elif chg >= 1 or rate >= 8.0:
            risk = "Moderate"
        else:
            risk = "Low Risk"

        routes_prone_to_charging.append({
            "route": r_name,
            "routeType": r_data["routeType"],
            "totalWeighed": tot,
            "chargedCount": chg,
            "warnedCount": r_data["warnedCount"],
            "legalCount": r_data["legalCount"],
            "chargeRate": rate,
            "riskLevel": risk,
            "topChargedCargos": top_cargos,
            "activeDaysCount": len(r_data["datesActive"]),
        })

    routes_prone_to_charging.sort(key=lambda x: (x["chargedCount"], x["chargeRate"], x["totalWeighed"]), reverse=True)

    # 7. Cross-Weighed Vehicles (Weighed on both Mobile and Static on the Same Day)
    static_records_by_date_plate: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)

    for s in month_sessions:
        is_static = (
            s.sections.get("daily_hour", {}).get("status") == "ready"
            or "overloaded" in s.dataframes
            or (storage_root / "processed" / s.report_id / "overloaded.pkl").exists()
        )
        if not is_static:
            continue

        for pkl_file in ["overloaded", "impounded_prohibited"]:
            sdf = s.dataframes.get(pkl_file)
            if sdf is None:
                sdf = _read_processed_pickle(storage_root, s.report_id, pkl_file)

            if sdf is not None and not sdf.empty:
                for _, row in sdf.iterrows():
                    reg = str(row.get("VehicleReg") or row.get("Registration") or "").strip().upper()
                    if not reg or reg in ["NAN", "NONE", "UNKNOWN", "-"]:
                        continue

                    dt_raw = str(row.get("DateWeighed") or row.get("Date Weighed/Prohibited") or "")
                    rec_date = dt_raw.split(" ")[0] if dt_raw and "-" in dt_raw else s.report_date
                    cargo_clean = _clean_cargo_name(row.get("Cargo"))

                    gvw_ov = float(row.get("GVWOverload") or 0)
                    axle_ov = float(row.get("AxleOverload") or 0)
                    status_raw = str(row.get("LastState") or row.get("Status") or row.get("state") or row.get("Vardict") or "Overloaded").strip()
                    st_eff = f"{status_raw} {row.get('Remarks', '')}".lower()

                    is_static_charged = (
                        "charg" in st_eff
                        or "court" in st_eff
                        or "prosecut" in st_eff
                        or pkl_file == "impounded_prohibited"
                        or max(gvw_ov, axle_ov) > 2000
                    ) and not any(k in st_eff for k in ["permit", "redistribut", "special release", "warned"])

                    if cargo_clean != "UNKNOWN" and is_static_charged and (gvw_ov > 0 or axle_ov > 0):
                        cargo_overload_counts[cargo_clean] += 1
                        cargo_overload_kg_total[cargo_clean] += max(gvw_ov, axle_ov)
                        if max(gvw_ov, axle_ov) > cargo_max_overload[cargo_clean]:
                            cargo_max_overload[cargo_clean] = max(gvw_ov, axle_ov)

                    static_records_by_date_plate[(rec_date, reg)].append({
                        "session_id": s.report_id,
                        "date": rec_date,
                        "time": dt_raw.split(" ")[1] if " " in dt_raw else "",
                        "registration": reg,
                        "station": s.station or s.weighbridge_name or "Static Weighbridge",
                        "bound": s.bound or "Static",
                        "cargo": cargo_clean,
                        "gvwOverloadKg": gvw_ov,
                        "axleOverloadKg": axle_ov,
                        "status": status_raw,
                        "ticketNo": str(row.get("TicketNo") or ""),
                    })

    cross_weighed_vehicles: list[dict[str, Any]] = []
    seen_cross_pairs: set[tuple[str, str, str, str]] = set()

    for (date_key, plate_key), m_list in mobile_records_by_date_plate.items():
        if (date_key, plate_key) in static_records_by_date_plate:
            s_list = static_records_by_date_plate[(date_key, plate_key)]
            for mr in m_list:
                for sr in s_list:
                    pair_id = (date_key, plate_key, mr["bound"], sr["bound"])
                    if pair_id in seen_cross_pairs:
                        continue
                    seen_cross_pairs.add(pair_id)

                    m_diff = mr["diffKg"]
                    s_axle = sr["axleOverloadKg"]
                    s_gvw = sr["gvwOverloadKg"]

                    if s_axle > 0 and m_diff <= 0:
                        perspective = (
                            f"Gross weight within legal tolerance on mobile patrol ({abs(m_diff):,.0f} kg margin), "
                            f"but flagged with {s_axle:,.0f} kg axle overload on static scale (axle load imbalance)."
                        )
                    elif s_axle > 0 and m_diff > 0:
                        perspective = (
                            f"Both mobile patrol and static weighbridge intercepted overload "
                            f"(Mobile diff +{m_diff:,.0f} kg, Static axle overload +{s_axle:,.0f} kg)."
                        )
                    elif s_gvw > 0:
                        perspective = (
                            f"Excess GVW detected at static facility (+{s_gvw:,.0f} kg overload) "
                            f"matching mobile field monitoring."
                        )
                    else:
                        perspective = (
                            f"Vehicle weighed on mobile patrol ({mr['remarks']}) and static facility ({sr['status']}) "
                            f"on the same calendar day."
                        )

                    cross_weighed_vehicles.append({
                        "date": date_key,
                        "registration": plate_key,
                        "cargo": mr["cargo"] if mr["cargo"] != "UNKNOWN" else sr["cargo"],
                        "mobile": {
                            "station": mr["station"],
                            "bound": mr["bound"],
                            "totalGvwKg": mr["totalGvwKg"],
                            "differenceKg": mr["diffKg"],
                            "remarks": mr["remarks"],
                            "origin": mr.get("origin"),
                            "destination": mr.get("destination"),
                        },
                        "static": {
                            "station": sr["station"],
                            "bound": sr["bound"],
                            "gvwOverloadKg": s_gvw,
                            "axleOverloadKg": s_axle,
                            "status": sr["status"],
                            "ticketNo": sr.get("ticketNo"),
                        },
                        "fieldPerspective": perspective,
                    })

    cross_weighed_vehicles.sort(key=lambda x: (x["date"], x["registration"]), reverse=True)

    # 8. Cargo Overloading Statistics (Vulnerability Ranking)
    total_overload_incidents = sum(cargo_overload_counts.values())
    cargo_overload_stats: list[dict[str, Any]] = []

    for cargo_name, count in cargo_overload_counts.most_common(15):
        tot_kg = cargo_overload_kg_total[cargo_name]
        avg_kg = round(tot_kg / count) if count > 0 else 0
        pct = round((count / total_overload_incidents * 100), 1) if total_overload_incidents > 0 else 0.0

        cargo_overload_stats.append({
            "cargo": cargo_name,
            "incidentCount": count,
            "totalExcessKg": round(tot_kg),
            "averageExcessKg": avg_kg,
            "maxExcessKg": round(cargo_max_overload[cargo_name]),
            "percentageShare": pct,
        })

    # 9. Return clean response with strict DMS/team omission
    return {
        "selectedMonth": target_month,
        "selectedMonthLabel": selected_month_label,
        "selectedMonthName": selected_month_name,
        "selectedYear": selected_year,
        "availableMonths": available_months_formatted,
        "station": target_code,
        "stationName": STATION_NAMES.get(target_code, target_code),
        "boundALabel": bound_a_name,
        "boundBLabel": bound_b_name,
        "kpis": {
            "totalTraffic": station_total_traffic,
            "thikaTraffic": station_bound_a_traffic,
            "nairobiTraffic": station_bound_b_traffic,
            "boundATraffic": station_bound_a_traffic,
            "boundBTraffic": station_bound_b_traffic,
            "totalCourtCases": station_total_cases,
            "thikaCourtCases": station_bound_a_cases,
            "nairobiCourtCases": station_bound_b_cases,
            "boundACourtCases": station_bound_a_cases,
            "boundBCourtCases": station_bound_b_cases,
            "complianceRate": round(station_compliance_rate, 1),
            "overloadsIntercepted": station_overloads_intercepted,
            "totalMobileWeighed": total_mobile_weighed,
            "totalMobileCharged": total_mobile_charged,
            "totalMobileWarned": total_mobile_warned,
            "totalMobileLegal": total_mobile_legal,
            "mobileChargeRate": round((total_mobile_charged / total_mobile_weighed * 100), 1) if total_mobile_weighed > 0 else 0.0,
        },
        "trafficData": traffic_data,
        "courtCasesData": court_cases_data,
        "crossStationData": cross_station_data,
        "routesProneToCharging": routes_prone_to_charging,
        "crossWeighedVehicles": cross_weighed_vehicles,
        "cargoOverloadStats": cargo_overload_stats,
    }
