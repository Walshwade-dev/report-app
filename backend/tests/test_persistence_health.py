from app import main


def test_persistence_health_fails_when_production_persistence_is_missing(
    client,
    temp_store,
    monkeypatch,
):
    monkeypatch.setenv("APP_ENV", "production")
    monkeypatch.delenv("REPORT_STORAGE_ROOT", raising=False)
    monkeypatch.setattr(main, "report_session_store", temp_store)
    monkeypatch.setattr(main, "SessionLocal", None)
    monkeypatch.setattr(main, "is_database_configured", lambda: False)

    response = client.get("/health/persistence")

    assert response.status_code == 503
    payload = response.json()
    assert payload["status"] == "error"
    assert payload["persistence_required"] is True
    assert payload["database"]["configured"] is False
    assert payload["database"]["connected"] is False
    assert payload["storage"]["configured"] is False


def test_persistence_health_allows_missing_persistence_in_development(
    client,
    temp_store,
    monkeypatch,
):
    monkeypatch.setenv("APP_ENV", "development")
    monkeypatch.delenv("REPORT_STORAGE_ROOT", raising=False)
    monkeypatch.setattr(main, "report_session_store", temp_store)
    monkeypatch.setattr(main, "SessionLocal", None)
    monkeypatch.setattr(main, "is_database_configured", lambda: False)

    response = client.get("/health/persistence")

    assert response.status_code == 200
    payload = response.json()
    assert payload["status"] == "ok"
    assert payload["persistence_required"] is False
    assert payload["database"]["configured"] is False
    assert payload["storage"]["configured"] is False
