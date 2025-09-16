import pandas as pd

# -------------------------------
# 1. โหลดข้อมูลจาก Excel
# -------------------------------
# อ่านไฟล์ Excel ทั้งหมด (ทุก sheet)
all_sheets = pd.read_excel("งานsupport.xlsx", sheet_name=None)

# แยกแต่ละ sheet ออกมาเป็น DataFrame
support1 = all_sheets.get("Support1")
support2 = all_sheets.get("Support2")
support3 = all_sheets.get("Support3")

# -------------------------------
# 2. ฟังก์ชันตรวจสอบ
# -------------------------------
def check_related(keyword):
    """
    ฟังก์ชันตรวจสอบความสัมพันธ์ของ keyword
    แสดงผลทั้ง:
      - ข้อมูลตัวเอง (Support1)
      - ความสัมพันธ์ 2 รายการ (Support2)
      - ความสัมพันธ์ 3 รายการ (Support3)
    """
    print(f"\n===== ตรวจสอบ: {keyword} =====")

    # --- Support1 (ตัวเอง) ---
    if support1 is not None:
        result1 = support1[support1["Item"] == keyword]
        if not result1.empty:
            row = result1.iloc[0]
            print(f"\n--- ข้อมูลตัวเอง ---")
            print(f"'{keyword}' พบ {row['Sum']} ครั้ง "
                  f"(Support={row['Support']*100:.1f}%)")

    # --- Support2 (ความสัมพันธ์ 2 รายการ) ---
    if support2 is not None:
        related2 = support2[
            (support2["Item1"] == keyword) | (support2["Item2"] == keyword)
        ].copy()

        if not related2.empty:
            # เรียงจาก Support มาก → น้อย
            related2 = related2.sort_values(by="Support", ascending=False)
            print("\n--- ความสัมพันธ์ 2 รายการ ---")
            for _, row in related2.iterrows():
                # ถ้า Item1 = keyword ให้ other = Item2, และกลับกัน
                other = row["Item1"] if row["Item2"] == keyword else row["Item2"]
                print(f"ถ้ามี '{keyword}' → มีโอกาสยื่น '{other}' "
                      f"(Support={row['Support']*100:.1f}%, "
                      f"Confidence={row['Cofidene']*100:.1f}%)")

    # --- Support3 (ความสัมพันธ์ 3 รายการ) ---
    if support3 is not None:
        related3 = support3[
            (support3["Item"] == keyword) |
            (support3["Item2"] == keyword) |
            (support3["Item3"] == keyword)
        ].copy()

        if not related3.empty:
            # เรียงจาก Support มาก → น้อย
            related3 = related3.sort_values(by="Support", ascending=False)
            print("\n--- ความสัมพันธ์ 3 รายการ ---")
            for _, row in related3.iterrows():
                # ดึงรายการอื่น ๆ ที่ไม่ใช่ keyword
                others = [x for x in [row["Item"], row["Item2"], row["Item3"]] if x != keyword]
                print(f"ถ้ามี '{keyword}' → มีโอกาสยื่น {', '.join(others)} "
                      f"(Support={row['Support']*100:.1f}%, "
                      f"Confidence={row['Cofidene']*100:.1f}%)")

# -------------------------------
# 3. ตัวอย่างการใช้งาน
# -------------------------------
check_related("ซื้อหนังสือ")
check_related("เที่ยวไทย")
