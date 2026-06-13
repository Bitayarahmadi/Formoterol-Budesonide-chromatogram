# -*- coding: utf-8 -*-
"""
Created on Mon May 18 13:50:44 2026

@author: Nafas
"""

import os
import zipfile
import tempfile
import re
from datetime import datetime

# تنظیمات اولیه
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
template = os.path.join(desktop, "100% inj3.docx")

def unique_name():
    i = 1
    while True:
        path = os.path.join(desktop, f"updated_report_{i}.docx")
        if not os.path.exists(path): return path
        i += 1

output = unique_name()

# ===== INPUT =====
sample = input("Sample Name: ")
inj = input("Injection Number: ")
date_input = input("Date Time (مثال 18.12.2025 09.00): ").strip()

dt = datetime.strptime(date_input, "%d.%m.%Y %H.%M")
date_word = dt.strftime("%m/%d/%Y %I:%M:%S %p")        # فرمت Injection Date
path_date = dt.strftime("%Y-%m-%d %H-%M-%S")          # فرمت مسیر فایل

# محاسبات
formo_pct = float(input("Formoterol %: "))
bude_pct = float(input("Budesonide %: "))

FORMO_AREA100, BUDE_AREA100 = 23.19004, 192.20139
FORMO_HEIGHT100, BUDE_HEIGHT100 = 4.85268, 24.56923
FORMO_WIDTH100, BUDE_WIDTH100 = 0.0796, 0.1304
FORMO_RT, BUDE_RT = "1.525", "3.179"

# تعریف متغیرهایی که باعث خطا شده بودند
formo_area, bude_area = FORMO_AREA100 * formo_pct / 100, BUDE_AREA100 * bude_pct / 100
formo_height, bude_height = FORMO_HEIGHT100 * formo_pct / 100, BUDE_HEIGHT100 * bude_pct / 100
formo_width, bude_width = FORMO_WIDTH100 * formo_pct / 100, BUDE_WIDTH100 * bude_pct / 100

total_area = formo_area + bude_area
total_height = formo_height + bude_height
formo_area_pct = (formo_area / total_area * 100) if total_area else 0
bude_area_pct = (bude_area / total_area * 100) if total_area else 0

# متون جدید
new_formo_row = f"1   {FORMO_RT} MM    {formo_width:.4f}   {formo_area:.5f}    {formo_height:.5f}  {formo_area_pct:.4f}"
new_bude_row = f"2   {BUDE_RT} MM    {bude_width:.4f}   {bude_area:.5f}   {bude_height:.5f}  {bude_area_pct:.4f}"
new_totals = f"Totals :                   {total_area:.5f}   {total_height:.5f}"

# ===== XML REWRITE =====
with tempfile.TemporaryDirectory() as tmp:
    with zipfile.ZipFile(template, "r") as z: z.extractall(tmp)

    for root_dir, _, files in os.walk(tmp):
        for file in files:
            if not file.endswith(".xml"): continue
            file_path = os.path.join(root_dir, file)
            with open(file_path, "r", encoding="utf-8") as f: content = f.read()

            # 1. جایگزینی متون ثابت
            content = re.sub(r"Sample Name:.*?100% 1404\.04\.01", f"Sample Name: {sample}", content)
            content = re.sub(r"1\s+1\.525\s+MM.*?(?=<)", new_formo_row, content)
            content = re.sub(r"2\s+3\.179\s+MM.*?(?=<)", new_bude_row, content)
            content = re.sub(r"Totals\s+:.*?(?=<)", new_totals, content)

            # 2. جایگزینی تاریخ‌ها
            content = re.sub(r'\d{2}(<[^>]+>)*/(<[^>]+>)*\d{2}(<[^>]+>)*/(<[^>]+>)*\d{4}.*?\d{2}.*?:.*?\d{2}.*?:.*?\d{2}.*?(AM|PM)', date_word, content)
            content = re.sub(r'\d{4}(<[^>]+>)*-(<[^>]+>)*\d{2}(<[^>]+>)*-(<[^>]+>)*\d{2}.*?\d{2}.*?-(<[^>]+>)*\d{2}.*?-(<[^>]+>)*\d{2}', path_date, content)

            with open(file_path, "w", encoding="utf-8") as f: f.write(content)

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as z:
        for root_dir, _, files in os.walk(tmp):
            for file in files:
                z.write(os.path.join(root_dir, file), os.path.relpath(os.path.join(root_dir, file), tmp))

print(f"✅ فایل نهایی ساخته شد: {output}")