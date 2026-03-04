import datetime
from datetime import timedelta
from spire.doc import *
from spire.doc.common import *
from flask import Flask, render_template, request, send_file, url_for
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.pagesizes import letter


#Defining existsInList function
def existsInList(element, alist):
    for e in alist:
        if e == element:
            return True
    return False

#Function for finding reading
def findReading(date):
    date_index = text1.find(date)
    if date_index == -1:
        return "---"
    else:
        paragraph_start_index = text1.find("-" ,date_index)
        paragraph_end_index = text1.find("*", date_index)
        return text1[paragraph_start_index + 1 : paragraph_end_index]

#Getting text from previous Morning Prayer Docs
inputdoc1 = Document()
inputdoc1.LoadFromFile("Morning Prayer Dates (Updated 3_4_2026).docx")
text1 = inputdoc1.GetText()

app = Flask(__name__)       

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate-pdf', methods=['GET', 'POST'])
def generate_pdf():
    if request.method == 'POST':
        # Retrieve user input from the form
        dateformat = '%Y-%m-%d'
        FirstDay = datetime.datetime.strptime(request.form['FirstDay'], dateformat)
        LastDay = datetime.datetime.strptime(request.form['LastDay'], dateformat)
        Easter = datetime.datetime.strptime(request.form['Easter'], dateformat)
        AshWed = Easter - timedelta(days = 46)

        # Adding Days Off
        days_off = []


        try:
            Break1Start = datetime.datetime.strptime(request.form['Break1Start'], dateformat)
            Break1End = datetime.datetime.strptime(request.form['Break1End'], dateformat)
        except:
            Break1Start = ''
            Break1End = ''
        while(Break1Start != '' and Break1Start <= Break1End):
            days_off.append(Break1Start)
            Break1Start = Break1Start + timedelta(days = 1)

        try:
            Break2Start = datetime.datetime.strptime(request.form['Break2Start'], dateformat)
            Break2End = datetime.datetime.strptime(request.form['Break2End'], dateformat)
        except:
            Break2Start = ''
            Break2End = ''
        while(Break2Start != '' and Break2Start <= Break2End):
            days_off.append(Break2Start)
            Break2Start = Break2Start + timedelta(days = 1)

        try:
            Break3Start = datetime.datetime.strptime(request.form['Break3Start'], dateformat)
            Break3End = datetime.datetime.strptime(request.form['Break3End'], dateformat)
        except:
            Break3Start = ''
            Break3End = ''
        while(Break3Start != '' and Break3Start <= Break3End):
            days_off.append(Break3Start)
            Break3Start = Break3Start + timedelta(days = 1)

        try:
            Break4Start = datetime.datetime.strptime(request.form['Break4Start'], dateformat)
            Break4End = datetime.datetime.strptime(request.form['Break4End'], dateformat)
        except:
            Break4Start = ''
            Break4End = ''
        while(Break4Start != '' and Break4Start <= Break4End):
            days_off.append(Break4Start)
            Break4Start = Break4Start + timedelta(days = 1)

        try:
            Break5Start = datetime.datetime.strptime(request.form['Break5Start'], dateformat)
            Break5End = datetime.datetime.strptime(request.form['Break5End'], dateformat)
        except:
            Break5Start = ''
            Break5End = ''
        while(Break5Start != '' and Break5Start <= Break5End):
            days_off.append(Break5Start)
            Break5Start = Break5Start + timedelta(days = 1)

        try:
            Break6Start = datetime.datetime.strptime(request.form['Break6Start'], dateformat)
            Break6End = datetime.datetime.strptime(request.form['Break6End'], dateformat)
        except:
            Break6Start = ''
            Break6End = ''
        while(Break6Start != '' and Break6Start <= Break6End):
            days_off.append(Break6Start)
            Break6Start = Break6Start + timedelta(days = 1)

        try:
            Break7Start = datetime.datetime.strptime(request.form['Break7Start'], dateformat)
            Break7End = datetime.datetime.strptime(request.form['Break7End'], dateformat)
        except:
            Break7Start = ''
            Break7End = ''
        while(Break7Start != '' and Break7Start <= Break7End):
            days_off.append(Break7Start)
            Break7Start = Break7Start + timedelta(days = 1)

        try:
            Break8Start = datetime.datetime.strptime(request.form['Break8Start'], dateformat)
            Break8End = datetime.datetime.strptime(request.form['Break8End'], dateformat)
        except:
            Break8Start = ''
            Break8End = ''
        while(Break8Start != '' and Break8Start <= Break8End):
            days_off.append(Break8Start)
            Break8Start = Break8Start + timedelta(days = 1)

        try:
            Break9Start = datetime.datetime.strptime(request.form['Break9Start'], dateformat)
            Break9End = datetime.datetime.strptime(request.form['Break9End'], dateformat)
        except:
            Break9Start = ''
            Break9End = ''
        while(Break9Start != '' and Break9Start <= Break9End):
            days_off.append(Break9Start)
            Break9Start = Break9Start + timedelta(days = 1)

        try:
            Break10Start = datetime.datetime.strptime(request.form['Break10Start'], dateformat)
            Break10End = datetime.datetime.strptime(request.form['Break10End'], dateformat)
        except:
            Break10Start = ''
            Break10End = ''
        while(Break10Start != '' and Break10Start <= Break10End):
            days_off.append(Break10Start)
            Break10Start = Break10Start + timedelta(days = 1)

    pdf_file = generate_pdf_file(FirstDay, LastDay, Easter, AshWed, days_off)
    return send_file(pdf_file, as_attachment=True, download_name=f"Morning Prayer {FirstDay.year}.pdf")

def generate_pdf_file(FirstDay, LastDay, easter, ash, daysoff):
    buffer = BytesIO()
    p = canvas.Canvas(buffer)
    width, height = letter
    
    #Printing  Result
    date = FirstDay

    while (date <= LastDay):
        if (date.strftime('%A') != "Saturday" and date.strftime('%A') != "Sunday" and existsInList(date, daysoff) == False):
            reading = findReading(str(date.month) + '/' + str(date.day))
            if date == (easter + timedelta(days = 9)):
                reading = findReading("Vocations1")
            elif (easter + timedelta(days = 9)) < date < (easter + timedelta(days = 21)):
                reading = findReading("Vocations2")
            elif date == (ash):
                reading = findReading("AshWed")
            elif date == (easter - timedelta(days = 6)):
                reading = findReading("HolyMon")
            elif date == (easter - timedelta(days = 5)):
                reading = findReading("HolyTues")
            elif date == (easter - timedelta(days = 4)):
                reading = findReading("HolyWed")
            elif date == (easter + timedelta(days = 2)):
                reading = findReading("OctTues")
            elif date == (easter + timedelta(days = 3)):
                reading = findReading("OctWed")
            elif date == (easter + timedelta(days = 4)):
                reading = findReading("OctThurs")
            elif date == (easter + timedelta(days = 5)):
                reading = findReading("OctFri")
            
            else:
                reading = findReading(str(date.month) + '/' + str(date.day))

            if date.strftime('%A') == "Friday" and (ash < date < easter):
                reading = reading + "\n" + findReading("LentFriday")

            printing_date = str(date.strftime("%A")) + ', ' + str(date.strftime("%B")) + ' ' + str(date.day) + ', ' + str(date.year)
            p1 = Paragraph(f"{printing_date} <br/><br/> Student 1 <br/><br/>Good morning.<br/>Please stand for our prayer and pledge of allegiance, and for all in the halls, please stop and pray.<br/><br/>My name is ___________ , and I am joined by  ___________.<br/><br/>Let us pray:<br/> In the name of the Father, and of the Son, and of the Holy Spirit.<br/><br/>{reading}<br/><br/>Pause<br/><br/>Student 2:<br/><br/>Caring for the needs of all here present let us pray together:<br/>Our Father….<br/>Hail Mary…<br/><br/>We offer these prayers and this day in the name of the Father and of the Son and of the Holy Spirit.  Amen.<br/><br/>And to honor our country, let us all say:<br/>I pledge allegiance to the flag…")
            p1.wrapOn(p, 450, 50)
            p1.drawOn(p, width-550, height-450)
            p.showPage()
            
        date = date + timedelta(days = 1)

    p.save()

    buffer.seek(0)
    return buffer

if __name__ == '__main__':
    app.run(debug=True)


