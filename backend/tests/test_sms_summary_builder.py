import pandas as pd
from app.services.report_session_store import ReportSession
from app.services.sms_summary_builder import build_static_sms_summary, build_mobile_sms_summary

def test_sms_summary_static():
    # Mocking ReportSession
    session = ReportSession(
        report_id="test_static_id",
        report_date="2026-06-25",
        station="Juja",
        bound="Thika Bound",
        weighbridge_name="Juja Weighbridge",
        prepared_by="ANASTASHA KENDA",
    )
    
    # Let's populate some daily_hour data
    # Columns: DATE, TIME, D, S, M, H, Q, X, C, Y, P, A, Z, G, R, E
    df_data = [
        ["2026-06-25", "0000-0100", 10, 20, 2, 100, 90, 32, 10, 5, 2, 3, 1, 1, 1, 0],
        ["Totals", "", 10, 20, 2, 100, 90, 32, 10, 5, 2, 3, 1, 1, 1, 0]
    ]
    df = pd.DataFrame(df_data, columns=["DATE", "TIME", "D", "S", "M", "H", "Q", "X", "C", "Y", "P", "A", "Z", "G", "R", "E"])
    
    session.dataframes["daily_hour"] = df
    session.sections["daily_hour"] = {"status": "ready"}
    session.manual_inputs["traffic_census"] = {"total_traffic_census": 500}
    
    summary = build_static_sms_summary(session)
    
    assert "JUJA THIKA BOUND WB" in summary
    assert "Date: 25.06.2026" in summary
    assert "Called in(C)=10" in summary
    assert "Weighed Hswim(Q)=90" in summary
    assert "T.traffic census(K)=500" in summary
    assert "T.traffic (T)=622" in summary # Q(90) + X(32) + K(500) + E(0) = 622
    assert "By: ANASTASHA KENDA." in summary


def test_sms_summary_mobile():
    session = ReportSession(
        report_id="test_mobile_id",
        report_date="2026-06-23",
        station="Juja mobile",
        bound="Mobile 2",
        weighbridge_name="Juja mobile",
        prepared_by="ANASTASHA KENDA",
    )
    
    # Columns in normalized mobile report
    # ticket_no, station, registration, deck_a_kg, deck_b_kg, deck_c_kg, deck_d_kg, gvw_kg, total_gvw_kg, gvw_difference_kg, remarks, is_weighed, is_dimension_charge, is_gvw_axle_charge, hour_band, mismatch, excess_kg
    df_data = [
        ["T1", "JJM2", "KDS042Z", 10000, 15000, 0, 0, 20000, 25000, 5000, "CHARGED", True, False, True, "0000-0100", "", 5000, pd.to_datetime("2026-06-23 00:30:00")],
        ["T2", "JJM2", "KBW781J", 5000, 10000, 0, 0, 18000, 15000, -3000, "LEGAL", True, False, False, "0100-0200", "", 0, pd.to_datetime("2026-06-23 01:30:00")]
    ]
    cols = ["ticket_no", "station", "registration", "deck_a_kg", "deck_b_kg", "deck_c_kg", "deck_d_kg", "gvw_kg", "total_gvw_kg", "gvw_difference_kg", "remarks", "is_weighed", "is_dimension_charge", "is_gvw_axle_charge", "hour_band", "mismatch", "excess_kg", "date_time"]
    df = pd.DataFrame(df_data, columns=cols)
    
    session.dataframes["mobile_report"] = df
    session.sections["mobile_report"] = {"status": "ready"}
    session.manual_inputs["mobile_report"] = {
        "route": "JUJA-KIMBO-RUIRU",
        "mileage_start": "100",
        "mileage_end": "250",
        "mobile_vehicle": "KDS042Z"
    }
    
    summary = build_mobile_sms_summary(session)
    
    assert "JUJA W/B DAILY MOBILE REPORT_TEAM TWO" in summary
    assert "Date:23.06.2026" in summary
    assert "Route:-JUJA-KIMBO-RUIRU" in summary
    assert "scale (S)=2" in summary
    assert "Total Weighed(X)=2" in summary
    assert "Total overloaded (Y)=1" in summary
    assert "Warned (A)=0" in summary
    assert "Legal=1" in summary
    assert "Charged & Prohibited(Z)=1" in summary
    assert "Above 2 Tonnes =1" in summary
    assert "Kilometers covered=150KMS" in summary
    assert "Vehicle used:- KDS042Z" in summary
    assert "By:-ANASTASHA KENDA." in summary


def test_sms_summary_metadata_fallback():
    session = ReportSession(
        report_id="test_static_id_fallback",
        report_date="2026-06-25",
        station="Juja",
        bound="Thika Bound",
        weighbridge_name="Juja Weighbridge",
        prepared_by="ANASTASHA KENDA",
    )
    
    # We do NOT populate session.dataframes! It is empty.
    # Instead, we populate the sections summary metadata.
    session.sections["daily_hour"] = {
        "status": "ready",
        "summary": {
            "D": 10, "S": 20, "M": 2, "H": 100, "Q": 90, "X": 32, "C": 10, "Y": 5, "P": 2, "A": 3, "Z": 1, "G": 1, "R": 1, "E": 0
        }
    }
    session.manual_inputs["traffic_census"] = {"total_traffic_census": 500}
    
    # Verify fallback total retrieval
    from app.services.sms_summary_builder import get_session_column_total
    assert get_session_column_total(session, "daily_hour", "Q") == 90
    assert get_session_column_total(session, "daily_hour", "X") == 32
    
    summary = build_static_sms_summary(session)
    
    assert "JUJA THIKA BOUND WB" in summary
    assert "Date: 25.06.2026" in summary
    assert "Called in(C)=10" in summary
    assert "Weighed Hswim(Q)=90" in summary
    assert "T.traffic census(K)=500" in summary
    assert "T.traffic (T)=622" in summary
    assert "By: ANASTASHA KENDA." in summary
