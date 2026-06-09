import webapp.app as webapp_module


def test_dashboard_responds_with_chart_placeholders() -> None:
    webapp_module.app.config["TESTING"] = True
    with webapp_module.app.test_client() as client:
        response = client.get("/")

    assert response.status_code == 200
    assert b'data-chart-key="status_distribution"' in response.data
    assert b'data-chart-key="impacted_amount_by_status"' in response.data
    assert b'data-chart-key="severity_distribution"' in response.data


def test_action_plan_responds_with_chart_placeholders() -> None:
    webapp_module.app.config["TESTING"] = True
    with webapp_module.app.test_client() as client:
        response = client.get("/action-plan")

    assert response.status_code == 200
    assert b'data-chart-key="top_customers_impacted_amount"' in response.data
    assert b'data-chart-key="top_customers_gap_count"' in response.data
    assert b'data-chart-key="severity_gap_distribution"' in response.data


def test_factures_absentes_route_responds_200() -> None:
    webapp_module.app.config["TESTING"] = True
    with webapp_module.app.test_client() as client:
        response = client.get("/factures-absentes")

    assert response.status_code == 200
    assert "Factures et avoirs absents RCI".encode("utf-8") in response.data


def test_historique_and_reference_quality_routes_respond_200() -> None:
    webapp_module.app.config["TESTING"] = True
    with webapp_module.app.test_client() as client:
        historique = client.get("/historique")
        reference_quality = client.get("/reference-quality")

    assert historique.status_code == 200
    assert b'data-chart-key="impacted_amount_trend"' in historique.data
    assert reference_quality.status_code == 200
    assert b'data-chart-key="top_out_scope_count"' in reference_quality.data


def test_chart_api_endpoints_respond_200() -> None:
    webapp_module.app.config["TESTING"] = True
    endpoints = [
        "/api/dashboard/charts",
        "/api/action-plan/charts",
        "/api/history/charts",
        "/api/reference-quality/charts",
    ]

    with webapp_module.app.test_client() as client:
        for endpoint in endpoints:
            response = client.get(endpoint)
            assert response.status_code == 200
            assert response.is_json
            assert isinstance(response.get_json(), dict)


def test_chart_helpers_return_clean_empty_data_when_files_are_absent(tmp_path, monkeypatch) -> None:
    empty_dirs = {
        "reports": tmp_path / "reports",
        "anomalies": tmp_path / "anomalies",
        "corrections": tmp_path / "corrections",
        "powerbi": tmp_path / "powerbi",
        "logs": tmp_path / "logs",
    }
    for key, value in empty_dirs.items():
        monkeypatch.setitem(webapp_module.DOWNLOAD_DIRS, key, value)

    assert webapp_module.load_latest_summary() == ({}, None)
    assert webapp_module.load_latest_reconciliation() == ([], None)
    assert webapp_module.load_reconciliation_history() == ([], None)
    assert webapp_module.load_latest_batch_control() == ([], None)
    assert webapp_module.load_reference_quality_rows() == ([], None)

    dashboard_charts = webapp_module.build_dashboard_charts([])
    action_plan_charts = webapp_module.build_action_plan_charts([])
    history_charts = webapp_module.build_history_charts([])
    reference_charts = webapp_module.build_reference_quality_charts([])

    assert dashboard_charts["status_distribution"]["empty"] is True
    assert action_plan_charts["top_customers_impacted_amount"]["empty"] is True
    assert history_charts["impacted_amount_trend"]["empty"] is True
    assert reference_charts["top_out_scope_count"]["empty"] is True
