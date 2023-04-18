import smtplib
import os
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#メール・SMTPサーバの設定
stmp_server ="smtp.gmail.com"
stmp_port = 587
stmp_user = "contens.host11@gmail.com"
stmp_password = "zjrdmbyfnxwmwsrw"
#宛先・件名・本文・ファイルパス・ファイル名の設定
to_address = sys.argv[1]
from_address = stmp_user
subject = "レポートの完成"
body = """
<html>
    <body>
        <h1>レポートが完成しました。</h1>
        <p>またのご利用をお待ちしております。</p>
    </body>
</html>"""
filepath = "image_file_storage/sales_forecast_report.xlsx"
filename = os.path.basename(filepath)

#メールメッセージを生成する
msg = MIMEMultipart()
msg["Subject"] = subject
msg["From"] = from_address
msg["To"] = to_address
msg.attach(MIMEText(body,"html"))
#メールにファイルを添付する
with open(filepath,"rb") as f:
    mb = MIMEApplication(f.read())
mb.add_header("Content-Disposition","attachment",filename=filename)
msg.attach(mb)

#SMTPサーバ接続・メール送信
s = smtplib.SMTP(stmp_server, stmp_port)
s.starttls()
s.login(stmp_user, stmp_password)
s.sendmail(from_address, to_address, msg.as_string())
s.quit()


