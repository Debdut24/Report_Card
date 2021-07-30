# Creator Name : Debdut Hansda
# Contact : hansdadev24@gmail.com
# Description : Application to create report cards from Excel data sheet and create them in PDFs
# Date of Creation : July 28, 2021

# ----------------------------------------- MODULES -----------------------------------------------------
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from win10toast import ToastNotifier

# -------------------------------------PDF CREATING FUNCTION---------------------------------------------
# FUNCTION TO MAKE REPORT CARD
def create_card(student_name):
    data = pd.read_excel ("Data/Dummy Data.xlsx")
    full_name = data[data["First Name "] == student_name]['Full Name '].unique ()[0]
    registration_no = data[data["First Name "] == student_name]['Registration Number'].unique ()[0]
    grade = data[data["First Name "] == student_name]['Grade '].unique ()[0]
    school_name = data[data["First Name "] == student_name]['Name of School '].unique ()[0]
    dob = data[data["First Name "] == student_name]['Date of Birth '].unique ()[0]
    dob = str (dob).split ("T")[0]
    city = data[data["First Name "] == student_name]['City of Residence'].unique ()[0]
    date_of_test = data[data["First Name "] == student_name]['Date and time of test'].unique ()[0]
    country = data[data["First Name "] == student_name]['Country of Residence'].unique ()[0]
    final_result = data[data["First Name "] == student_name]['Final result'].unique ()[0]
    max_marks = data[data["First Name "] == student_name]['Score if correct'].sum ()
    student_marks = data[data["First Name "] == student_name]['Your score'].sum ()

    filename = f"PDF/{full_name}_record.pdf"
    title = "Report Card"
    image = f"Pics/{full_name}.png"

    # MARKS INFO
    marks = data[data['Full Name '] == full_name]
    marks = marks[['Question No.', 'What you marked', 'Correct Answer',
                   'Outcome (Correct/Incorrect/Not Attempted)', 'Score if correct',
                   'Your score']]

    mark_dict = marks.to_dict ()
    mark_list = [['Question No.', 'What you marked', 'Correct Answer',
                  'Outcome (Correct/Incorrect/Not Attempted)', 'Score if correct',
                  'Your score']]

    for n in marks.index:
        mark_list.append ([mark_dict['Question No.'][n],
                           mark_dict['What you marked'][n],
                           mark_dict['Correct Answer'][n],
                           mark_dict['Outcome (Correct/Incorrect/Not Attempted)'][n],
                           mark_dict['Score if correct'][n],
                           mark_dict['Your score'][n]
                           ])

    # creating a pdf object
    pdf = canvas.Canvas (filename)

    # creating the info section
    text = pdf.beginText (5, 720)
    text.setFont ("Courier", 14)
    text.setFillColor (colors.black)
    text.textLine (f"Full Name : {full_name}")
    text.textLine (f"Registration No : {registration_no}")
    text.textLine (f"Grade : {str (grade)}")
    text.textLine (f"Name of school :{school_name}")
    text.textLine (f"Date of Birth : {dob}")
    text.textLine (f"City of Residence : {city}")
    text.textLine (f"Date and time of test : {date_of_test}")
    text.textLine (f"Country of Residence : {country}")
    pdf.drawText (text)

    # creating the heading section
    heading = pdf.beginText (20, 560)
    heading.setFont ("Courier-Bold", 16)
    heading.setFillColor (colors.fidred)
    heading.textLine (f"                 Marks Scored :{student_marks}/{max_marks}")
    heading.textLine (final_result)
    pdf.drawText (heading)

    # inserting logo
    pdf.drawImage ("atom.png", x = 250, y = 750, width = 80, height = 80)

    # inserting student image
    pdf.drawImage (image, x = 420, y = 600, width = 150, height = 170)

    # table creation
    t = Table (mark_list)

    # table styling
    style = TableStyle ([
        ('BACKGROUND', (0, 0), (-1, 0), colors.green),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('BOX', (0, 0), (-1, -1), 2, colors.black),
        ('GRID', (0, 1), (-1, -1), 1, colors.black)
    ])

    t.setStyle (style)
    for i in range (1, len (mark_list)):
        if i % 2 == 0:
            bc = colors.burlywood
        else:
            bc = colors.beige
        ts = TableStyle (
            [('BACKGROUND', (0, i), (-1, i), bc)]
        )
        t.setStyle (ts)

    width = 450
    height = 700
    x = 5
    y = 50
    t.wrapOn (pdf, width, height)
    t.drawOn (pdf, x, y)

    # creation of pdf
    pdf.save ()

# ----------------------------------------- MAIN BODY ------------------------------------------------
# GATHER INFO FROM THE DATA SHEET AND IMAGES OF THE STUDENTS
# stores the data from data sheet
data = pd.read_excel("Data/Dummy Data.xlsx")
# list of the students name
student_list = data['First Name '].unique()
total_marks_scored = data[data["First Name "]== "ABC1"]["Your score"].sum()
total_marks = data[data["First Name "]== "ABC1"]['Score if correct'].sum()
# list to store the info of each and every student as separate item
student_info = []
image_list = []
for student in student_list:
    create_card(student)
# ----------------------------------DISPLAY SUCCESSFUL MESSAGE----------------------------------
toast = ToastNotifier()
toast.show_toast("Python","Successful",duration = 1)