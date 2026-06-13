# -*- coding: utf-8 -*-
"""
Created on Sun May 17 15:14:31 2026

@author: Nafas
"""

import os
import zipfile
import tempfile
from datetime import datetime
import xml.etree.ElementTree as ET
import re

# ------------------------------
# مسیر فایل ورودی
# ------------------------------
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
template = os.path.join(desktop, "100% inj3.docx")   # فایل الگو

# ------------------------------
# ساخت نام یکتا در دسکتاپ
# ------------------------------
def unique_name():
    i = 1
    while True:
        path = os.path.join(desktop, f"updated_report_{i}.docx")
        if not os.path.exists(path):
            return path
        i += 1

output = unique_name()

# ------------------------------
# ورودی‌ها
# ------------------------------
sample = input("Sample Name: ").strip()
inj = input("Injection Number: ").strip()
date_input = input("Date Time (مثال 18.09.2017 01.21): ").strip()

dt = datetime.strptime(date_input, "%d.%m.%Y %H.%M")
date_word = dt.strftime("%m/%d/%Y %I:%M:%S %p")   # نمایش کامل
date_short1 = dt.strftime("%d-%b-%y")             # کوتاه
date_short2 = dt.strftime("%Y-%m-%d %H-%M-%S")    # نوع فایل‌ها

formo_pct = float(input("Formoterol %: "))
bude_pct = float(input("Budesonide %: "))

# ------------------------------
# مقادیر ثابت
# ------------------------------
FORMO_AREA100 = 23.19004
BUDE_AREA100 = 192.20139
FORMO_HEIGHT100 = 4.85268
BUDE_HEIGHT100 = 24.56923
FORMO_WIDTH100 = 0.0796
BUDE_WIDTH100 = 0.1304
FORMO_RT = "1.525"
BUDE_RT = "3.179"

# ------------------------------
# محاسبات
# ------------------------------
formo_area = FORMO_AREA100 * formo_pct / 100
bude_area = BUDE_AREA100 * bude_pct / 100
formo_height = FORMO_HEIGHT100 * formo_pct / 100
bude_height = BUDE_HEIGHT100 * bude_pct / 100
formo_width = FORMO_WIDTH100 * formo_pct / 100
bude_width = BUDE_WIDTH100 * bude_pct / 100

total_area = formo_area + bude_area
total_height = formo_height + bude_height
formo_area_pct = formo_area / total_area * 100 if total_area else 0
bude_area_pct = bude_area / total_area * 100 if total_area else 0

new_formo_row = f"1   {FORMO_RT} MM    {formo_width:.4f}   {formo_area:.5f}    {formo_height:.5f}  {formo_area_pct:.4f}"
new_bude_row = f"2   {BUDE_RT} MM    {bude_width:.4f}   {bude_area:.5f}    {bude_height:.5f}  {bude_area_pct:.4f}"
new_totals = f"Totals :                   {total_area:.5f}   {total_height:.5f}"

# پیشوندهای قابل شناسایی
old_sample_prefix = "Sample Name:"
old_date_prefix = "Injection Date  :"
old_formo_row_prefix = "1   1.525 MM"
old_bude_row_prefix = "2   3.179 MM"
old_totals_prefix = "Totals :"

# regex برای هر نوع تاریخ
date_patterns = [
    re.compile(r"\b\d{2}[./-]\d{2}[./-]\d{4}\b"),                   # 18.09.2017 یا 18-09-2017
    re.compile(r"\b\d{4}[./-]\d{2}[./-]\d{2}\b"),                   # 2017-09-18
    re.compile(r"\b\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+(AM|PM)\b"),  # 09/18/2017 01:21:00 PM
    re.compile(r"\b\d{4}-\d{2}-\d{2}\s+\d{2}-\d{2}-\d{2}\b"),
    re.compile(r"\b\d{2}-[A-Za-z]{3}-\d{2,4}\b")
]

# ------------------------------
# ویرایش XML درون docx
# ------------------------------
with tempfile.TemporaryDirectory() as tmp:
    with zipfile.ZipFile(template, "r") as z:
        z.extractall(tmp)

    for root_dir, _, files in os.walk(tmp):
        for file in files:
            if not file.endswith(".xml"):
                continue

            file_path = os.path.join(root_dir, file)
            tree = ET.parse(file_path)
            root = tree.getroot()

            modified = False

            for elem in root.iter():
                if not elem.tag.endswith("}t") or elem.text is None:
                    continue
                text = elem.text

                # --- Sample Name ---
                if old_sample_prefix in text:
                    elem.text = f"{old_sample_prefix} {sample}"
                    modified = True

                # --- Injection Date line ---
                elif old_date_prefix in text:
                    elem.text = f"{old_date_prefix} {date_word}"
                    modified = True

                # --- Chromatogram rows ---
                elif text.strip().startswith(old_formo_row_prefix):
                    elem.text = new_formo_row
                    modified = True
                elif text.strip().startswith(old_bude_row_prefix):
                    elem.text = new_bude_row
                    modified = True
                elif text.strip().startswith(old_totals_prefix):
                    elem.text = new_totals
                    modified = True

                # --- Generic date replacement everywhere ---
                else:
                    new_text = text
                    for pattern in date_patterns:
                        new_text = pattern.sub(date_word, new_text)
                    if new_text != text:
                        elem.text = new_text
                        modified = True

            if modified:
                tree.write(file_path, encoding="utf-8", xml_declaration=True)

    # بازسازی docx
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as z:
        for root_dir, _, files in os.walk(tmp):
            for file in files:
                full_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(full_path, tmp)
                z.write(full_path, rel_path)

print("✅ فایل جدید ذخیره شد:", output)
print("✅ Sample:", sample)
print("✅ Injection:", inj)
print("✅ Date:", date_word)
print("✅ Formoterol row:", new_formo_row)
print("✅ Budesonide row:", new_bude_row)
print("✅ Totals:", new_totals)
