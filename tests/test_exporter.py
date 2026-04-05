# tests/test_exporter.py
import sys
import os

# Добавляем корень проекта в путь, чтобы импорты работали
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from modules.schema import ACCERecord
from modules.exporter import export_to_acce


def create_mock_data():
    records = []

    pump = ACCERecord()
    pump.source_file = "Samples 4.xlsx"
    pump.raw_tag = "423-P-001A"
    pump.user_tag = "423-P-001A"
    pump.description_ru = "Насос для оборотной воды"
    pump.quantity = 1
    pump.price_code = "6852"
    pump.acce_item_symbol = "CP"
    pump.acce_item_type = "CENTRIF"
    pump.is_valid = True
    pump.parent_area = "0407.423"

    pump.flow_rate = 160.0
    pump.flow_rate_unit = "м3/ч"
    pump.pressure = 0.8
    pump.pressure_unit = "МПа"
    pump.motor_power_kw = 22.0
    records.append(pump)

    fan = ACCERecord()
    fan.source_file = "Samples 4.xlsx"
    fan.raw_tag = "423-D-101"
    fan.user_tag = "423-D-101"
    fan.description_ru = "Вентилятор"
    fan.quantity = 1
    fan.price_code = "6852"
    fan.acce_item_symbol = "FN"
    fan.acce_item_type = "CENTRIF"
    fan.is_valid = True
    fan.parent_area = "0407.423"

    fan.flow_rate = 5000.0
    fan.pressure = 2500.0  # Па
    fan.pressure_unit = "Па"
    fan.motor_power_kw = 15.0
    records.append(fan)

    boiler = ACCERecord()
    boiler.source_file = "Samples 4.xlsx"
    boiler.raw_tag = "423-B-001A"
    boiler.user_tag = "423-B-001A"
    boiler.description_ru = "Котел ВОТ"
    boiler.quantity = 1
    boiler.price_code = "6880.12"
    boiler.acce_item_symbol = "STB"
    boiler.acce_item_type = "PACKAGE"
    boiler.is_valid = True
    boiler.parent_area = "0407.423"

    boiler.capacity_kw = 4000.0
    boiler.pressure = 0.8
    boiler.pressure_unit = "МПа"
    records.append(boiler)

    # 4. Реактор (VT) - из Пример 3.docx (Р-001)
    reactor = ACCERecord()
    reactor.source_file = "Пример 3.docx"
    reactor.raw_tag = "Р-001"
    reactor.user_tag = "R-001"  # Транслитерация
    reactor.description_ru = "Реактор окисления"
    reactor.quantity = 1
    reactor.acce_item_symbol = "VT"
    reactor.acce_item_type = "CYLINDER"
    reactor.is_valid = True
    reactor.parent_area = "UNIT_REACTORS"

    reactor.diameter_m = 0.45
    reactor.length_m = 3.95
    reactor.design_temperature = 450.0
    reactor.pressure = 4.9
    reactor.pressure_unit = "МПа"
    reactor.material = "12ХМ"
    records.append(reactor)

    # 5. Невалидная запись (для проверки листа VALIDATION)
    bad_rec = ACCERecord()
    bad_rec.raw_tag = "BAD-TAG"
    bad_rec.user_tag = "BAD-TAG"
    bad_rec.description_ru = "Ошибка"
    bad_rec.is_valid = False
    bad_rec.errors = ["Missing Item Symbol"]
    records.append(bad_rec)

    return records


if __name__ == "__main__":
    print("🛠 Generating mock data...")
    mock_records = create_mock_data()

    # Пути
    # Шаблон будет создан автоматически функцией _create_dummy_template, если его нет
    template_file = "tests/dummy_acce_template.xlsx"
    output_file = "output/test_export_result.xlsx"

    # Убедимся, что папка output существует
    os.makedirs("output", exist_ok=True)
    os.makedirs("tests", exist_ok=True)

    print(f"🚀 Exporting {len([r for r in mock_records if r.is_valid])} valid records...")

    try:
        result = export_to_acce(
            records=mock_records,
            template_path=template_file,
            output_path=output_file,
            project_name="Test Project 2026"
        )

        print(f"\n✅ SUCCESS!")
        print(f"File saved to: {result.output_path}")
        print(f"Exported: {result.exported_count} items")
        print(f"Skipped: {result.skipped_count} items")
        print(f"By Type: {result.by_type}")
        if result.warnings:
            print(f"Warnings: {result.warnings}")

        print("\n📋 Next steps:")
        print("1. Open the file in Excel.")
        print("2. Check sheets: Contents, AREAS, CP, FN, STB, VT, VALIDATION.")
        print("3. Verify that 'BAD-TAG' is in VALIDATION sheet with status INVALID.")
        print("4. Verify that pressures are converted to kPa (e.g., 0.8 MPa -> 800 kPa).")

    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback

        traceback.print_exc()