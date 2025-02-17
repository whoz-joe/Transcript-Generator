import csv
import pandas as pd
import math
import openpyxl 
import os
from pywebio import *
from pywebio.input import *
from pywebio.output import *
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment,Font, colors,Border,Side
from openpyxl.drawing.image import Image
import smtplib, ssl
from email.message import EmailMessage


roll_to_email = {}

def send_individual_mail(server, file , email) :
    print(file , email)
    msg = EmailMessage()
    msg['Subject'] = "Quiz marks CSE"
    msg['From'] = "Shivam Singh Kushwah and Meghana Reddy"
    msg['To'] = email[0] + "," + email[1]

    with open("content.txt") as content :
        msg_content = content.read()
        msg.set_content(msg_content)
    


    with open("sample_output/marksheet/" + file , "rb") as f:
        file_data = f.read()
        # print(file_name)
        msg.add_attachment(file_data, maintype="application" ,subtype="xlsx" ,filename = file)
    # print(msg)
    try:
        server.send_message(msg)
    except :
        print("Some error occured , " , email , file)
    print("email sent to " , email)


def send_email() :
    print("Sending email in progress")
    put_text("Sending Emails to students.......")
    server = smtplib.SMTP_SSL('smtp.gmail.com',465)
    server.login('danielisgamingtoday@gmail.com' , "thisIsPassword@")
    for file in os.listdir("sample_output/marksheet") :
        if file == "concise_marksheet.csv" :
            continue 
        send_individual_mail(server, file,roll_to_email[file])
    put_success("Emails sent to concerned students")



def marksheetfromroll(roll,Name,ans,stud_ans,total,neg_marks,pos,neg):
    wb = openpyxl.Workbook()
    total_qns = len(ans)
    img = Image("logo.png")
    sheet = wb.active
    sheet.add_image(img, "A1")
    sheet.title = "Quiz.marks"

    sheet.column_dimensions['A'].width = 25
    sheet.column_dimensions['B'].width = 25
    sheet.column_dimensions['C'].width = 25
    sheet.column_dimensions['D'].width = 25
    sheet.column_dimensions['E'].width = 25

    b_font = Font(bold=True,size=18)
    n_font = Font(size=14)
    green_font= Font(bold=True,size=14,color = '00008000')
    red_font= Font(bold=True,size=14,color = '00FF0000')
    blue_font= Font(bold=True,size=14,color = '000000FF')
    
    sheet.merge_cells('A6:E7')
    font = Font(size=17,bold=True,underline='single')
    cell = sheet.cell(row=6, column=1)  
    cell.value = "marksheet"
    cell.font = font 
    cell.alignment = Alignment(horizontal='center', vertical='center')  

    cell = sheet["A8"]
    cell.value = "Name"
    cell.font = n_font

    cell = sheet["B8"]
    cell.value = Name
    cell.font = b_font

    cell = sheet["D8"]
    cell.value = "Exam"
    cell.font = n_font

    cell = sheet["E8"]
    cell.value = "Quiz"
    cell.font = n_font

    cell = sheet["A9"]
    cell.value = "Roll Number"
    cell.font = n_font

    cell = sheet["B9"]
    cell.value = roll
    cell.font = b_font

    double = Side(border_style='medium', color="00000000")
    border = Border(left=double,right=double,top=double,bottom=double)
    cell = sheet["A11"]
    cell.border = border

    cell = sheet["B11"]
    cell.value = "Right"
    cell.font = b_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["C11"]
    cell.value = "Wrong"
    cell.font = b_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["D11"]
    cell.value = "Not Attempt"
    cell.font = b_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["E11"]
    cell.value = "max"
    cell.font = b_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center') 

    cell = sheet["A14"]
    cell.border = border
    cell.value = "No."
    cell.font = b_font
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["B14"]
    cell.value = total/pos
    cell.font = green_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["C14"]
    cell.value = neg_marks/neg
    cell.font = red_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["D14"]
    cell.value = total_qns - (total/pos + neg_marks/neg)
    cell.font = b_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["E14"]
    cell.value = total_qns
    cell.font = blue_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center') 


    cell = sheet["A13"]
    cell.border = border
    cell.value = "marking"
    cell.font = b_font
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["B13"]
    cell.value = pos
    cell.font = green_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["C13"]
    cell.value = neg
    cell.font = red_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["D13"]
    cell.value = 0
    cell.font = b_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["E13"]
    cell.font = blue_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center') 


    cell = sheet["A14"]
    cell.border = border
    cell.value = "Total"
    cell.font = b_font
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["B14"]
    cell.value = total
    cell.font = green_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["C14"]
    cell.value = neg_marks
    cell.font = red_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["D14"]
    cell.font = b_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center')  

    cell = sheet["E14"]
    cell.value = str(total+neg_marks) + "/" + str(pos*total_qns)
    cell.font = blue_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center') 

    cell = sheet["A17"]
    cell.value = "Student Ans"
    cell.font = b_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center') 

    cell = sheet["B17"]
    cell.value = "Correct Ans"
    cell.font = b_font
    cell.border = border
    cell.alignment = Alignment(horizontal='center') 

    for i in range(1,total_qns+1):

        cell = sheet["A"+str(17+i)]
        if isinstance(stud_ans[i],str):
            cell.value = stud_ans[i]
        elif isinstance(stud_ans[i], float) and math.isnan(stud_ans[i]):
            cell.value = ""
        
        if stud_ans[i] == ans[i]:
            cell.font = green_font
        else:
            cell.font = red_font
        cell.border = border
        cell.alignment = Alignment(horizontal='center') 

        cell = sheet["B"+str(17+i)]
        cell.value = ans[i]
        cell.font = blue_font
        cell.border = border
        cell.alignment = Alignment(horizontal='center') 
    wb.save("sample_output/marksheet/"+roll+'.xlsx')

def marksheets(csvreader,masreader,pos,neg):
    stdmpp = {}
    header = []
    for ind in masreader.index:
        stdmpp[masreader['roll'][ind]] = masreader['name'][ind]

    with open("sample_input/save_response.csv","w+") as csvfile:
        csvreader.to_csv("sample_input/save_response.csv")

    df = pd.read_csv("sample_input/save_response.csv")
    print(df)

    score_after_negative = []
    statusAns = []

    ans = getAnswermapFromRoll("ANSWER", df)
    print(ans)
    for ind in df.index:
        stud_ans = getAnswermapFromRoll(df['Roll Number'][ind], df) 
        roll_to_email[df['Roll Number'][ind] + ".xlsx"]=[df['Email address'][ind] , df["IITP webmail"][ind] ]
        total = 0
        actual_marks = 0
        neg_marks = 0
        for i in range(1,len(stud_ans)+1):
            if stud_ans[i] == ans[i]:
                total += pos
            elif isinstance(stud_ans[i], float) and math.isnan(stud_ans[i]):
                pass
            else:
                neg_marks +=neg 
        actual_marks = total + neg_marks
        score_after_negative.append(str(actual_marks)+"/"+str(pos*len(stud_ans)))
        statusAns.append("["+str(int(total/pos))+","+str(int(neg_marks/neg))+","+str(int(len(stud_ans)-total/pos-neg_marks/neg))+"]")
        marksheetfromroll(df['Roll Number'][ind], df['Name'][ind],ans,stud_ans,total,neg_marks,pos,neg)


    print(len(stud_ans))
    df.insert(loc=6, column='Score_After_Negative', value=score_after_negative)
    df.insert(loc=len(df.columns), column='statusAns', value=statusAns)
    df.to_csv('sample_output/marksheet/concise_marksheet.csv',index=False)

def main():
    # put_html("<h1> Please select the master Roll file : </h1>")
    master_roll = file_upload(label='Please select the master Roll file', accept=".csv", name=None, placeholder='Choose file', multiple=False, max_size=0, max_total_size=0, required=True)
    master_roll_csv = content_to_pandas(master_roll['content'].decode('utf-8').splitlines())
    # print(master_roll_csv)
    # print(master_roll['content'])
    response_file = file_upload(label='Please select the Responses file', accept=".csv", name=None, placeholder='Choose file', multiple=False, max_size=0, max_total_size=0, required=True)
    response_file_csv = content_to_pandas(response_file['content'].decode('utf-8').splitlines())
    # print(response_file_csv)
    positive_marks = input("What is the positive marks per question" , type = NUMBER)

    neg_marks = input("What is negative marks per question (write 0 if none) " , type = NUMBER)

    # master_roll_csv = read_csv(master_roll)
    # response_file_csv = read_csv(response_file)
    # print(master_roll_csv)
    # print(response_file_csv)

    marksheets(response_file_csv , master_roll_csv, positive_marks, neg_marks)

    put_success("Yayy!!! Your marksheets have been generated", closable=True, scope=None, position=- 1) 

    req_action = "generate pdf"
    while(req_action != "none"):
        req_action = actions(label="Please Select your action", 
                    buttons=[{'label': 'Download marksheets', 'value': "marksheets"}, 
                             {'label':'Send Email to students', 'value': "Email"} , 
                             {'label':'I wish to exit', 'value': "none"}
                             ])
        if(req_action == "marksheets"):
            put_html("<h2>Your marksheets have been downloaded, please check your sample_output folder</h2>")
        if(req_action == "Email") : 
            send_email()
            put_html("<h2>The mail have been sent to the concered students</h2>")

    put_success("This was our project, hope you liked it :)")

def content_to_pandas(content: list):
    with open("tmp.csv", "w") as csv_file:
        writer = csv.writer(csv_file, delimiter =",")
        for line in content:
            # print(line)
            writer.writerow(line.split(","))
    return pd.read_csv("tmp.csv")

def getAnswermapFromRoll(roll,df):
    ans = df.loc[df['Roll Number'] == roll]
    answers = ans.values.tolist()[0][7:]
    answer_map = {}
    c =1
    for ans in answers:
        answer_map[c] = ans
        c+=1 
    return answer_map
    
start_server(main, port=3001, debug=True , static_dir = "/")