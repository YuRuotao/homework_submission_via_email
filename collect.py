from email.header import Header
from email.mime.text import MIMEText
import email
import smtplib
import imaplib
import json
from base64 import b64decode
import os
#basic set up
PASS, MAIL, IMAPSERVER, SMTPSERVER, SIGN = None,None,None,None,None,
with open("config.json","r",encoding="utf-8") as c:
    config=json.loads(c.read())
    PASS = config["PASS"]
    MAIL = config["MAIL"]
    IMAPSERVER = config["IMAPSERVER"]
    SMTPSERVER = config["SMTPSERVER"]
    SIGN = config["SIGN"]

print("MAIL: ",MAIL)
print("IMAPSERVER: ",IMAPSERVER)
print("SMTPSERVER: ",SMTPSERVER)

#log in
conn = imaplib.IMAP4_SSL(IMAPSERVER, "993")  # ssl
imaplib.Commands['ID'] = ('AUTH')
smtpObj = smtplib.SMTP()
try:
    conn.login(MAIL, PASS)
    smtpObj.connect(SMTPSERVER, 25)
    smtpObj.login(MAIL, PASS)
    args = ("name", "gongchuang201", "contact", "gongchuang201@163.com", "version", "1.0.0", "vendor", "myclient")
    typ, dat = conn._simple_command('ID', '("' + '" "'.join(args) + '")')
except Exception as e:
    print('Error: %s' % str(e.args[0], "utf-8"))
    exit()
print("CONNECTED!")

#get email
print("Fetching mail list")
conn.select()
typ, data = conn.search(None, 'ALL')

download_folder = input("download folder, for example E:/school_work/study_representative/gongchuang201/homework:")
workid = input("Today's Work's ID:")
reply_check = int(input("to send reply, enter 1:"))

#Get Student List
students = []
with open("student.txt", "r", encoding="utf-8") as f:
    students = f.readlines()

students = map(lambda x: x.rstrip("\n").split("\t"), students)  # 按照"学号 名字"处理
students = map(lambda x: [x[0], [x[1], False]], students)       # 添加是否提交的判断
students = dict(students)

newlist = data[0].split()[::-1]  # 列表化收到的邮件，时间倒序
for i in newlist:
    print("Fetching mail %d" % int(i))
    typ, data = conn.fetch(i, '(RFC822)')
    print("Analysing mail %d" % int(i))
    msg = email.message_from_string(data[0][1].decode('utf-8'))
    sub = email.header.decode_header(msg.get('subject'))[0][0]  # 标题
    if type(sub) == bytes:
        sub = sub.decode('utf-8')
    sub_0 = sub.split("/")  # 标题格式 学号/作业号
    if len(sub_0) == 2:
        if sub_0[1] != workid:  # 判断是不是本次作业
            continue
    else:
        continue
    pass
    stu = sub_0[0]  # 前半截是学号
    print("[SUBJECT] ", sub)
    mailfrom = email.utils.parseaddr(msg.get('from'))[1]  # 发件人,回信要用
    print("[   FROM] ", mailfrom)
    foo = 0  # 多文件防重名

# saving files
    if msg.get_content_maintype() == 'multipart':
    # loop on the parts of the mail
        for part in msg.walk():
    #find the attachment part - so skip all the other parts
            if part.get_content_maintype() == 'multipart': continue
            if part.get_content_maintype() == 'text': continue
            if part.get('Content-Disposition') == 'inline': continue
            if part.get('Content-Disposition') is None: continue
            name = part.get_filename()
            break
    # fix filename
        if name[0:8] == "=?utf-8?":
            name = name[10:-2]
            name = b64decode(name).decode(encoding="utf-8")
        if name:
            print("[ ATTACH] ", name)
            attach_data = part.get_payload(decode=True)  # 附件数据
            filename = stu + students[stu][0]

            # saving attachments
            #E:/school_work/study_representative/gongchuang201/homework
            att_path = os.path.join(download_folder, filename)
            if not os.path.isfile(att_path):
                fp = open(att_path, 'wb')
                fp.write(attach_data)
                fp.close()
                foo += 1
        else:# 不是附件
            pass
    students[stu][1] = True  # 标记提交

# 发送回信
    if reply_check == 1:
        mail_msg = "<p>" + \
            students[stu][0] + \
            "同学：<br>你的作业"+workid + \
            "已接收。<br>祝生活愉快！</p><hr>"+SIGN
        message = email.mime.text.MIMEText(mail_msg, 'html', 'utf-8')
        message['Subject'] = Header('Re:'+sub, 'utf-8')
        message['From'] = Header(MAIL, 'utf-8')
        message['To'] = Header(students[stu][0], 'utf-8')
        try:
            smtpObj.sendmail(MAIL, [mailfrom], message.as_string())
            print("Reply Sent")
        except smtplib.SMTPException:
            print("Error: Reply not sent")

# 输出提交情况
for i, j in students.items():
    print(i, "|", j[0], "\t|", "submitted" if j[1] else "NOT submitted")

conn.logout()
