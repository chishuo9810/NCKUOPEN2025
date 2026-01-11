from flask import Flask, request, render_template, send_from_directory
import logging
from datetime import datetime
from pathlib import Path
import os
import pandas as pd

app = Flask(__name__)


# 建立/設定 logger

Path("logs").mkdir(exist_ok=True)  # logs 資料夾
logging.basicConfig(
    filename=f"logs/data.log",
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

@app.after_request
def log_request(resp):
    if resp.status_code == 304:
        return resp
    logging.info(
        '%s %s "%s %s" %s "%s"',
        request.remote_addr,
        request.headers.get("Host"),
        request.method,
        request.full_path,
        resp.status_code,
        request.headers.get("User-Agent")
    )
    return resp

# 讀 Excel
file_path = "./2025check_in.xlsx"
df = pd.read_excel(file_path, engine="openpyxl")

# ACME 驗證路由
@app.route("/.well-known/acme-challenge/<path:filename>")
def acme_challenge(filename):
    return send_from_directory(
        os.path.join(app.root_path, ".well-known", "acme-challenge"),
        filename
    )

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        name = request.form["name"]
        if name:
            try:
                result = df.loc[df["選手"] == name]
                if not result.empty:
                    player_number   = int(result["選手編號"].values[0])
                    player_name     = result["選手"].values[0]
                    team_name       = result["隊名"].values[0] if pd.notna(result["隊名"].values[0]) else "❌無"
                    size            = result["衣服尺寸"].values[0]
                    first           = result["項目1"].values[0]
                    # main_group      = result["主要組別"].values[0]
                    # classify_number = result["分類碼"].values[0]
                    second          = result["項目2"].values[0] if pd.notna(result["項目2"].values[0]) else "❌無"
                    # font_size       = 20 if len(player_name) < 5 else 30

                    if len(result) == 1:
                        return render_template(
                            "index.html",
                            player_name=player_name,
                            player_number=player_number,
                            size=size,
                            # font_size=font_size,
                            first=first,
                            second=second,
                            team_name=team_name,
                        )
                    else: 
                        player_number2   = int(result["選手編號"].values[1])
                        player_name2    = result["選手"].values[1]
                        team_name2       = result["隊名"].values[1] if pd.notna(result["隊名"].values[1]) else "❌無"
                        size2            = result["衣服尺寸"].values[1]
                        first2           = result["項目1"].values[1]
                        # main_group2      = result["主要組別"].values[1]
                        # classify_number2 = result["分類碼"].values[1]
                        second2          = result["項目2"].values[1] if pd.notna(result["項目2"].values[1]) else "❌無"
                        # font_size2       = 20 if len(player_name2) < 5 else 30
                        return render_template(
                            "index.html",
                            player_name=player_name,
                            player_number=player_number,
                            size=size,
                            # font_size=font_size,
                            first=first,
                            second=second,
                            team_name=team_name,
                            player_name2=player_name2,
                            player_number2=player_number2,
                            size2=size2,
                            # font_size2=font_size2,
                            first2=first2,
                            second2=second2,
                            team_name2=team_name2,
                        )
                else:
                    return render_template("index.html", error="❌找不到對應的選手訊息!")
            except (IndexError, ValueError):
                return render_template("index.html", error="❌系統錯誤")
    return render_template("index.html")


# 啟動 (HTTPS)

if __name__ == "__main__":
    app.run(
        # host="0.0.0.0",
        # port=443,
        # ssl_context=("nckuopen.duckdns.org-crt.pem", "nckuopen.duckdns.org-key.pem"),
        debug=True
    )
