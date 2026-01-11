from flask import Flask, request, render_template, redirect, url_for
import pandas as pd
import os

app = Flask(__name__)

# 讀取Excel文件
file_path = './2025nckuopen.xlsx'
df = pd.read_excel(file_path, engine='openpyxl')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        name = request.form['name']
        if name:
            try:
                result = df.loc[df['選手'] == name]
                if not result.empty:
                    player_number = int(result['選手編號'].values[0])
                    player_name = result['選手'].values[0]
                    team_name = result['隊名'].values[0]
                    size = result['衣服尺寸'].values[0]
                    first = result['項目1'].values[0]
                    main_group = result['主要組別'].values[0]
                    classify_number = result['分類碼'].values[0]
                    second = result['項目2'].values[0] if pd.notna(result['項目2'].values[0]) else "無項目2"
                    font_size = 20 if len(player_name) < 5 else 30

                    return render_template(
                        'index.html',
                        player_name=player_name,
                        player_number=player_number,
                        size=size,
                        font_size=font_size,
                        first=first,
                        second=second,
                        team_name=team_name
                    )
                else:
                    return render_template('index.html', error="找不到對應的選手資訊")
            except (IndexError, ValueError):
                return render_template('index.html', error="資料處理錯誤，請檢查Excel內容")
    return render_template('index.html')

if __name__ == '__main__':
    # 注意：請將 cert.pem / key.pem 憑證放在同一個資料夾下
    ssl_cert = 'cert.pem'
    ssl_key = 'key.pem'

    if not os.path.exists(ssl_cert) or not os.path.exists(ssl_key):
        print("❌ 找不到憑證檔案 cert.pem / key.pem，請先用 Win-ACME 申請並放在此資料夾")
        exit(1)

    app.run(host='0.0.0.0', port=443, ssl_context=(ssl_cert, ssl_key))
