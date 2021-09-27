import smtplib
from email.mime.text import MIMEText
from email.header import Header


# 发送邮件，需要第三方的smtp服务器，这里的密码是在邮箱网站申请授权码，不是自己的登录密码
mail_host = "smtp.qq.com"
mail_sender = '101341517@qq.com'
mail_pass = 'vqinuerkoyhhbiad'  # 授权码

# 接收邮件
mail_receivers = ['newync@hotmail.com']

# 邮件内容，文本格式，把plain改成html是html格式
message = MIMEText('邮件内容 -- HelloWorld', 'plain', 'utf-8')
# 显示发件人
message['From'] = Header(mail_sender)
# 显示收件人
message['To'] = ','.join(mail_receivers)
message['Subject'] = Header('Python邮件测试', 'utf-8')

try:
    # QQ邮箱SMTP服务器smtp.qq.com（端口465或587）SSL一般用465
    smtpObj = smtplib.SMTP_SSL(mail_host, 465)
    smtpObj.login(mail_sender, mail_pass)
    smtpObj.sendmail(mail_sender, message['To'].split(','), message.as_string())
    smtpObj.quit()
    print('邮件发送成功')
except smtplib.SMTPException:
    print("Error: 发送失败")
