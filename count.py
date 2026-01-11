from flask import Flask, request, render_template, send_from_directory
import logging
from datetime import datetime
from pathlib import Path
import os
import pandas as pd
from collections import Counter


# 讀 Excel
file_path = "./2025nckuopen.xlsx"
df = pd.read_excel(file_path, engine="openpyxl")

name_counts = Counter(df["選手"])

with open("a.txt", "w", encoding="utf-8") as f:
    for name, count in name_counts.items():
        if count != 1:
            f.write(f"{name}{count}\n")    # 出現多次就加數字（如：王曉明2）
