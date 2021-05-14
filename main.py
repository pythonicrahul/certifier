from flask_mail import Mail
import os
import json
from flask import Flask,render_template,request,session, redirect
from flask_sqlalchemy import SQLAlchemy
import xlrd
from PIL import Image, ImageDraw, ImageFont
from werkzeug.utils import secure_filename
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
with open("config.json", "r") as f:
    params = json.load(f)["params"]

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = params["upload"]
app.config.update(
    MAIL_SERVER='smtp.gmail.com',
    MAIL_PORT='465',
    MAIL_USE_SSL=True,
    MAIL_USERNAME=params["email"],
    MAIL_PASSWORD=params["password"]
)

MAIL_USERNAME=params["email"]
MAIL_PASSWORD=params["password"]
mail = Mail(app)


def collect_Data(loc):
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    i = 1
    participants = []
    while i < sheet.nrows:
        name = sheet.row_values(i)[1]
        year = sheet.row_values(i)[2]
        branch = sheet.row_values(i)[3]
        event = sheet.row_values(i)[4]
        position = sheet.row_values(i)[5]
        email = sheet.row_values(i)[6]
        l = [name, year, branch, event, position,email]
        participants.append(l)
        # image.save(f'send{sno}.png')
        i = i + 1
    return participants


def generate_certificate(name, year, branch, event, position, email, sno):
    sno = str(sno)
    # position = str(int(position))
    date = params["date"]
    fontTitle = ImageFont.truetype('title.ttf', size=70)
    fontName = ImageFont.truetype('title.ttf', size=100)
    image = Image.open('blank_certificate.jpg')
    draw = ImageDraw.Draw(image)

    color = 'rgb(0, 0, 0)'
    color2 = 'rgb(255,0,0)'
    (x, y) = (1300, 1370)
    draw.text((x, y), name, fill=color, font=fontName, spacing=6)
    (x, y) = (625, 1605)
    draw.text((x, y), year, fill=color, font=fontTitle, spacing=6)
    (x, y) = (1175, 1605)
    draw.text((x, y), branch, fill=color, font=fontTitle, spacing=6)
    (x, y) = (630, 1705)
    draw.text((x, y), event, fill=color, font=fontTitle, spacing=6)
    (x, y) = (455, 785)
    draw.text((x,y), sno , fill=color2, font=fontTitle, spacing=6)
    (x, y) = (2877, 1600)
    draw.text((x, y), position, fill=color, font=fontTitle, spacing=6)
    (x, y) = (1300, 1790)
    draw.text((x, y), date, fill=color, font=fontTitle, spacing=6)
    image.save(f'{name}.png')
    body = params["body_of_email"]
    msg = MIMEMultipart()
    msg['From'] = params["email"]
    msg['To'] = email
    msg['Subject'] = ' Congratulations you got it ... '
    msg.attach(MIMEText(body, 'plain'))
    filename = f'{name}.png'
    attachment = open(filename, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= " + filename)
    msg.attach(part)
    text = msg.as_string()
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(MAIL_USERNAME, MAIL_PASSWORD)
    server.sendmail(MAIL_USERNAME, email, text)
    server.quit()


excel_filename = ""


@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        global excel_filename
        f = request.files['upl']
        print(f.filename)
        f.save(os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f.filename)))
        excel_filename = f.filename
        return redirect("/waiting")
    return render_template("index.html")


@app.route("/waiting")
def waiting():
    global excel_filename
    loc = app.config['UPLOAD_FOLDER'] + "\\" + excel_filename
    data = collect_Data(loc)
    sno = 1
    f = open("data.txt", "a")

    for row in data:
        sno = sno + 1
        generate_certificate(row[0], row[1], row[2], row[3], row[4], row[5], sno)
        f.write(row[0] + " " + row[2] + " " + str(sno) + "\n")
    f.close()
    return render_template("waiting.html")


if __name__ == "__main__":
    
    app.run(debug=True)
