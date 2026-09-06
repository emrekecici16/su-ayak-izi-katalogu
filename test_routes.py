"""Flask route/uç nokta testleri."""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import suayakizi  # noqa: E402

import pytest  # noqa: E402


@pytest.fixture
def client():
    suayakizi.app.config["TESTING"] = True
    with suayakizi.app.test_client() as c:
        yield c


def test_index_returns_200(client):
    r = client.get("/")
    assert r.status_code == 200
    assert "Su Ayak İzi".encode("utf-8") in r.data


def test_index_has_no_unreplaced_placeholder(client):
    r = client.get("/")
    assert b"__SUAYAK_DATA_JSON__" not in r.data


def test_index_defaults_to_light_theme(client):
    r = client.get("/")
    assert b'data-theme="light"' in r.data


def test_index_has_no_long_dash_characters(client):
    r = client.get("/")
    text = r.data.decode("utf-8")
    assert "\u2014" not in text, "Sayfada uzun tire (\u2014) tespit edildi."
    assert "\u2013" not in text, "Sayfada kısa çizgi (\u2013) tespit edildi."


def test_manifest(client):
    r = client.get("/manifest.json")
    assert r.status_code == 200
    assert r.mimetype == "application/manifest+json"


def test_service_worker(client):
    r = client.get("/sw.js")
    assert r.status_code == 200


def test_icon_found(client):
    r = client.get("/icons/icon-192.png")
    assert r.status_code == 200
    assert r.mimetype == "image/png"


def test_maskable_icon_found(client):
    r = client.get("/icons/icon-512-maskable.png")
    assert r.status_code == 200


def test_icon_not_found(client):
    r = client.get("/icons/does-not-exist.png")
    assert r.status_code == 404


def test_404_page(client):
    r = client.get("/bu-sayfa-yok")
    assert r.status_code == 404
    assert b"404" in r.data


def test_api_sectors_all(client):
    r = client.get("/api/sectors")
    data = json.loads(r.data)
    assert data["count"] == data["total"] == len(suayakizi.DATA)


def test_api_sectors_filter_by_sector(client):
    sector = suayakizi.SECTORS[0]
    r = client.get("/api/sectors", query_string={"sector": sector})
    data = json.loads(r.data)
    assert data["count"] > 0
    assert all(row["Sektör"] == sector for row in data["results"])


def test_api_sectors_filter_by_type(client):
    r = client.get("/api/sectors", query_string={"type": "Mavi Su"})
    data = json.loads(r.data)
    assert data["count"] > 0
    assert all(row["Su Ayak İzi Türü"] == "Mavi Su" for row in data["results"])


def test_api_sectors_search(client):
    r = client.get("/api/sectors", query_string={"q": "tekstil"})
    data = json.loads(r.data)
    assert data["count"] > 0


def test_export_xlsx_all(client):
    r = client.get("/export/xlsx")
    assert r.status_code == 200
    assert r.mimetype == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert r.data[:2] == b"PK"
    assert "su-ayak-izi-katalog.xlsx" in r.headers["Content-Disposition"]


def test_export_xlsx_filtered(client):
    sector = suayakizi.SECTORS[0]
    r = client.get("/export/xlsx", query_string={"sector": sector, "type": "Karma"})
    assert r.status_code == 200
    assert r.data[:2] == b"PK"
