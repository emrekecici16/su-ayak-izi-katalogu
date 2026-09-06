"""su_ayak_izi_sektorleri.json veri bütünlüğü testleri."""
import json
import os

DATA_FILE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "su_ayak_izi_sektorleri.json"
)
REQUIRED_FIELDS = (
    "Sektör", "Süreç", "Su Yoğunluğu", "Su Ayak İzi Türü",
    "Başlıca Su Riski", "İlgili Standartlar", "Azaltım Uygulamaları", "Açıklama",
)
KNOWN_INTENSITY = {"Düşük", "Orta", "Yüksek", "Çok Yüksek"}
KNOWN_TYPES = {"Mavi Su", "Yeşil Su", "Gri Su", "Karma"}


def load_raw():
    with open(DATA_FILE, encoding="utf-8") as f:
        return json.load(f)


def test_data_file_exists():
    assert os.path.isfile(DATA_FILE)


def test_data_is_nonempty_list():
    data = load_raw()
    assert isinstance(data, list)
    assert len(data) > 0


def test_every_record_has_required_fields():
    data = load_raw()
    for i, row in enumerate(data):
        for field in REQUIRED_FIELDS:
            assert row.get(field), f"Kayıt #{i} alanı eksik: {field}"


def test_no_duplicate_process_names_within_sector():
    data = load_raw()
    keys = [(row["Sektör"], row["Süreç"]) for row in data]
    duplicates = {k for k in keys if keys.count(k) > 1}
    assert not duplicates, f"Aynı sektörde yinelenen süreç adı: {duplicates}"


def test_intensity_values_are_known():
    data = load_raw()
    for row in data:
        assert row["Su Yoğunluğu"] in KNOWN_INTENSITY, (
            f"Bilinmeyen su yoğunluğu: {row['Su Yoğunluğu']} (Süreç: {row['Süreç']})"
        )


def test_type_values_are_known():
    data = load_raw()
    for row in data:
        assert row["Su Ayak İzi Türü"] in KNOWN_TYPES, (
            f"Bilinmeyen su ayak izi türü: {row['Su Ayak İzi Türü']} (Süreç: {row['Süreç']})"
        )


def test_every_sector_label_has_expected_format():
    data = load_raw()
    for row in data:
        assert row["Sektör"][0].isdigit(), f"Beklenmeyen sektör formatı: {row['Sektör']}"


def test_no_long_dash_characters_anywhere_in_data():
    """Veri setinde uzun tire (em dash) veya kısa çizgi (en dash) bulunmamalı."""
    raw_text = json.dumps(load_raw(), ensure_ascii=False)
    assert "\u2014" not in raw_text, "Veri setinde uzun tire (\u2014) tespit edildi."
    assert "\u2013" not in raw_text, "Veri setinde kısa çizgi (\u2013) tespit edildi."
