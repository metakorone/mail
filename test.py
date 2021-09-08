import smtplib
from email.mime.text import MIMEText

# 送受信先
to_addr = "sakauekanade@gmail.com; yidongxingfeng@gmail.com"
from_addr = "ito.y.bs@m.titech.ac.jp"

msg = MIMEText('本文', "plain", 'utf-8')
msg['Subject'] = "メールのタイトル"
msg['From'] = from_addr
msg['To'] = to_addr

with smtplib.SMTP_SSL(host="smtpv3.m.titech.ac.jp", port=465) as smtp:
    smtp.login('ito.y.bs$njtpp5', '1224Kanade')
    smtp.send_message(msg)
    smtp.quit()