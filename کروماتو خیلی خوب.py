# -*- coding: utf-8 -*-
"""
Created on Wed May  6 15:37:18 2026

@author: Nafas
"""

import os
import zipfile
import tempfile
from datetime import datetime
import xml.etree.ElementTree as ET

desktop = os.path.join(os.path.expanduser("~"), "Desktop")
template = os.path.join(desktop, "100% inj3.docx")

def unique_name():
    i = 1
    while True:
        path = os.path.join(desktop, f"updated_report_{i}.docx")
        if not os.path.exists(path):
            return path
        i += 1

output = unique_name()

# ===== INPUT =====
sample = input("Sample Name: ")
inj = input("Injection Number: ")
date_input = input("Date Time (مثال 18.09.2017 01.21): ")


# *** رشته‌ای که می‌خواهیم حذفش کنیم ***
old_bude_numbers_block = "192.20139   24.56923  89.2335"

# ===== XML REWRITE =====
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

            for elem in root.iter():
                if not elem.tag.endswith("}t") or elem.text is None:
                    continue

                text = elem.text

                if old_sample in text:
                    elem.text = text.replace(old_sample, new_sample)

                if old_date_prefix in text and ("PM" in text or "AM" in text):
                    elem.text = new_date

                if old_formo_row_prefix in text:
                    elem.text = new_formo_row

                if old_bude_row_prefix in text:
                    elem.text = new_bude_row

                if old_totals_prefix in text:
                    elem.text = new_totals

                # *** این قسمت جدید است: حذف بلوک بودزوناید قدیمی در هر جا که باشد ***
                if old_bude_numbers_block in text:
                    elem.text = text.replace(old_bude_numbers_block, "")

            tree.write(file_path, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as z:
        for root_dir, _, files in os.walk(tmp):
            for file in files:
                full_path = os.path.join(root_dir, file)
                rel = os.path.relpath(full_path, tmp)
                z.write(full_path, rel)

print("✅ Saved:", output)
print("✅ New Sample:", sample)
print("✅ New Date:", date_word)
print("✅ New Formoterol row:", new_formo_row)
print("✅ New Budesonide row:", new_bude_row)
print("✅ New Totals:", new_totals)
