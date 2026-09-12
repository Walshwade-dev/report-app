from app.services.analytics_service import get_analytics_details


def test_analytics_details_structure_and_defaults():
    res = get_analytics_details(station="Juja")
    assert "selectedMonth" in res
    assert "selectedMonthLabel" in res
    assert "availableMonths" in res
    assert isinstance(res["availableMonths"], list)
    assert "kpis" in res
    assert "routesProneToCharging" in res
    assert "crossWeighedVehicles" in res
    assert "cargoOverloadStats" in res

    # Verify KPI keys
    kpis = res["kpis"]
    assert "totalTraffic" in kpis
    assert "complianceRate" in kpis
    assert "totalMobileWeighed" in kpis
    assert "totalMobileCharged" in kpis


def test_analytics_details_specific_month_cross_weighing():
    # 2026-06 has cross-weighed vehicles between mobile and static
    res = get_analytics_details(station="Juja", month="2026-06")
    assert res["selectedMonth"] == "2026-06"

    # Verify cross-weighed vehicles
    cross = res["crossWeighedVehicles"]
    assert len(cross) > 0
    sample = cross[0]
    assert "registration" in sample
    assert "date" in sample
    assert "mobile" in sample
    assert "static" in sample
    assert "fieldPerspective" in sample

    # Verify routes prone to charging
    routes = res["routesProneToCharging"]
    assert isinstance(routes, list)
    if routes:
        r = routes[0]
        assert "route" in r
        assert "totalWeighed" in r
        assert "chargedCount" in r
        assert "chargeRate" in r

    # Verify cargo overload stats
    cargos = res["cargoOverloadStats"]
    assert isinstance(cargos, list)
    if cargos:
        c = cargos[0]
        assert "cargo" in c
        assert "incidentCount" in c
        assert "totalExcessKg" in c
        assert "averageExcessKg" in c


def test_no_dms_name_or_team_in_analytics():
    res = get_analytics_details(station="Juja", month="2026-06")

    # Serialize to string and verify no DMS / team keys or names
    import json
    payload_str = json.dumps(res).lower()

    assert "dms_name" not in payload_str
    assert "dm entry" not in payload_str
    assert "danka_staff" not in payload_str
    assert "police_officers" not in payload_str
