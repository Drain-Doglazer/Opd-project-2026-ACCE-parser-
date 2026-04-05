import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules.schema import ACCERecord
from modules.exporter import export_to_excel


def main():
    records = []

    pump = ACCERecord()
    pump.source_file = "test_manual"
    pump.raw_tag = "P-101A"
    pump.user_tag = "P-101A"
    pump.description_ru = "Центробежный насос подачи воды"
    pump.quantity = 1
    pump.acce_item_symbol = "CP"
    pump.is_valid = True

    pump.flow_rate = 150.0
    pump.motor_power_kw = 45.0
    pump.design_temperature = 25.0
    pump.pressure = 1.2  # МПа

    records.append(pump)

    vessel = ACCERecord()
    vessel.source_file = "test_manual"
    vessel.raw_tag = "V-201"
    vessel.user_tag = "V-201"
    vessel.description_ru = "Сепаратор газа вертикальный"
    vessel.parent_area = "UNIT_200"
    vessel.quantity = 1
    vessel.acce_item_symbol = "VT"
    vessel.is_valid = True
    vessel.material = "Carbon Steel"

    vessel.volume = 50.0
    vessel.diameter_m = 2.5
    vessel.length_m = 12.0
    vessel.design_temperature = 150.0
    vessel.pressure = 4.5

    records.append(vessel)

    bad_record = ACCERecord()
    bad_record.raw_tag = "P-999"
    bad_record.user_tag = "P-999"
    bad_record.description_ru = "Сломанный насос"
    bad_record.acce_item_symbol = "CP"
    bad_record.is_valid = False
    records.append(bad_record)

    output_file = "output/test_acce_export.xlsx"

    os.makedirs("output", exist_ok=True)

    print(f"Exporting {len([r for r in records if r.is_valid])} valid records...")
    try:
        export_to_excel(records, output_file)
        print(f"Success! File saved to: {output_file}")
        print("Check the file in Excel. It should have sheets: Contents, AREAS, PUMPS, VESSELS.")
        print("P-999 should NOT be present.")
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()