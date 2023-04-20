from flask import Flask
from flask import request
from flask import render_template
import subprocess
import shutil
import os

app = Flask(__name__)
#HTMLテンプレートファイルを読み込む
@app.route("/")
def main_menu():
    return render_template("index.html")

@app.route("/report",methods=["POST"])
def create_report():
    #データ・ファイルを受け取る
    file = request.files["file"]
    file_name = file.filename
    email = request.form["email"]
    report_type = request.form.get("report")

    #受け取ったデータ・ファイルの確認
    if email == "":
        return render_template("index.html",message = "メールアドレスを入力してください!")
    if file_name == "":
        return render_template("index.html",message = "Excelファイルを選択してください!")
    if report_type == None:
        return render_template("index.html",message = "レポートの種類が選択されていません!")
    
    os.makedirs("image_file_storage", exist_ok=True)
    file.save("image_file_storage/send_file.xlsx")

    #レポート種類を判別
    if report_type == "A":
        subprocess.run(["python", "./static/py/sales_forecast_report.py"])
    if report_type == "B":
        subprocess.run(["python", "./static/py/summary_report.py"])
    
    #メール送信ファイル実行
    subprocess.run(["python", "./static/py/send_email.py",email])
    shutil.rmtree("image_file_storage")
    
    #レポートの作成が終わったことを報告
    return render_template("index.html",message="ご利用ありがとうございます、レポートの作成が完了したことをお知らせいたします。")

if __name__ == "__main__":
    # port = int(os.getenv("PORT", 5000))
    # app.run(host="0.0.0.0", port=port)
    app.debug = True
    app.run(host='localhost')