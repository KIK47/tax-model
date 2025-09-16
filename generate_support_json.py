# -------------------------------
# generate_support_json.py
# -------------------------------
# สคริปต์นี้ใช้แปลงไฟล์ Excel "งานsupport.xlsx"
# เป็นไฟล์ JSON "support_data.json" ที่ HTML/JS จะโหลดไปแสดงผล
# -------------------------------

import pandas as pd
import json

# -------------------------------
# 1. โหลดข้อมูลจาก Excel
# -------------------------------
# ใช้ pandas.read_excel อ่านไฟล์ทั้งหมด
excel_file = "งานsupport.xlsx"
all_sheets = pd.read_excel(excel_file, sheet_name=None)  # sheet_name=None = โหลดทุก sheet

# แยก sheet ออกเป็นตัวแปร
support1 = all_sheets.get("Support1")  # ตาราง Support1
support2 = all_sheets.get("Support2")  # ตาราง Support2
support3 = all_sheets.get("Support3")  # ตาราง Support3

# -------------------------------
# 2. แปลงข้อมูลเป็น JSON
# -------------------------------
# to_dict(orient="records") = แปลง DataFrame เป็น list ของ dict
data_json = {
    "support1": support1.to_dict(orient="records") if support1 is not None else [],
    "support2": support2.to_dict(orient="records") if support2 is not None else [],
    "support3": support3.to_dict(orient="records") if support3 is not None else []
}

# -------------------------------
# 3. บันทึกเป็นไฟล์ JSON
# -------------------------------
output_file = "support_data.json"
with open(output_file, "w", encoding="utf-8") as f:
    json.dump(data_json, f, ensure_ascii=False, indent=2)  # ensure_ascii=False เพื่อเก็บภาษาไทย, indent=2 ให้ดูง่าย

print(f"แปลงไฟล์ Excel → JSON เรียบร้อย: {output_file}")
