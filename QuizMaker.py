import csv
import pygame
import gspread
from gspread import *
from gspread import utils
from gspread_formatting import *
from oauth2client.service_account import ServiceAccountCredentials
import random
import time

pygame.init()

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)

client = gspread.authorize(creds)

print("\nWelcome to Mr. Sajches Pixel Art Google Sheets Quiz Generator")

backslash = '\\'
while True:
        sheetName = input("\nPlease enter the name of the Google Sheets to be created: ")
        if backslash in sheetName:
            print("\nInvalid Name for Google Sheet, please try again")
        else:
            break

imageAddress = input("\nEnter the file location of the image: ")

while True:
    try:
        image = pygame.image.load(str(imageAddress))
        print('file {} found!'.format(imageAddress))
        break
    except FileNotFoundError:
        print('\n \nWoah, we couldnt find that image, please make sure you copy and paste the location of the image.\n  (You can put the image in the same folder as the exe and type the filename)')
        imageAddress = input("\nEnter the exact file location of the image: ")

try:
    sheet = client.open(sheetName)
    print('\nspreadsheet found, replacing spreadsheet')
    sheets = client.del_spreadsheet(sheet.id)
    sheet = client.create(sheetName)
    sheet2 = sheet.add_worksheet("Image", rows=image.get_height(), cols=image.get_width())
except SpreadsheetNotFound:
    print('\nspreadsheet not found, creating spreadsheet')
    sheet = client.create(sheetName)
    sheet2 = sheet.add_worksheet("Image", rows=image.get_height(), cols=image.get_width())

quizMakerRunning = True
questions = []
answers = []
while quizMakerRunning:
    print("\nWill you be loading a quiz file from a csv, or will you create your own quiz?")
    load = input("""
    1. Create a quiz
    2. Load Quiz from .csv file

    Enter Value: """)
    if load == '2':
        while quizMakerRunning:
            csvName = input("Please enter the address of the csv file you would like to import: ")
            try:
                with open(str(csvName), 'r') as csvQuiz:
                    csvReader = csv.reader(csvQuiz)
                    for line in csvReader:
                        print(line)
                        questions.append(line[0])
                        answers.append(line[1])
                        print(questions)
                        print(answers)
                quizMakerRunning = False
            except:
                print("csv file not found, please enter the exact address of the file.\n\nYou can put the file in the same folder as the exe and enter the file name.\n")
    elif load == '1':
        print("""
        Please enter your questions and answers

        -Use the format: QUESTION:ANSWER
        -Example --> is the moon made of cheese?:no
        -Do not put digits as answers, the pixels will not show up if you do
        -Please dont put spaces before or after the colon
        -If you want to start over type RESET
        -If you want to delete your previous entry, type BACK
        -To redo a specific entry, type REDO
        -To remove an entry, type REMOVE
        -To view all your entrys, type LIST

        -When you are done, type DONE
        """)
        while quizMakerRunning:
            try:
                entry = input('\nQuestion {}: '.format(len(questions)+1))
                if entry.__contains__(':') and entry.count(':') == 1:
                    print(entry)
                    entry = entry.split(':')
                    questions.append(entry[0])
                    answers.append(entry[1])
                elif entry == 'BACK':
                    if len(questions)>0:
                        questions.pop()
                        answers.pop()
                        print('\nQuestion {} removed'.format(len(questions)+1))
                    else:
                        print("No questions or answers inputted")
                elif entry == 'RESET':
                    questions.clear()
                    answers.clear()
                    print("Quiz Reset\n")
                elif entry == 'LIST':
                    print('\n#. QUESTION:ANSWER')
                    for index in list(range(len(questions))):
                        print('{}. {}:{}'.format(index+1, questions[index], answers[index]))
                elif entry == 'REDO':
                    #list the questions like in LIST but then promt a number to change, then promt again with question number, then replace question and answer with new entry
                    print('\n#. QUESTION:ANSWER')
                    for index in list(range(len(questions))):
                        print('{}. {}:{}'.format(index+1, questions[index], answers[index]))
                    num = int(input("\nEnter the index of the question you want to replace: "))
                    if isinstance(num, int) and 0 < num <= (int(len(questions))):
                        replacement = input("Re-enter the question and answer: ")
                        if replacement.__contains__(':') and replacement.count(':') == 1:
                            replacement = replacement.split(':')
                            questions[num-1] = replacement[0]
                            answers[num-1] = replacement[1]
                            print('\nAnswer replaced')
                        else:
                            raise ValueError      
                    else:
                        print("Number out of range or incorrect format, please try again")
                elif entry == 'DONE':
                    if len(questions) > 0:
                        quizMakerRunning = False
                    else:
                        print("Must have at least one question!")
                elif entry == 'REMOVE':
                    print('\n#. QUESTION:ANSWER')
                    for index in list(range(len(questions))):
                        print('{}. {}:{}'.format(index+1, questions[index], answers[index]))
                    num = int(input("\nEnter the index of the question you want to remove: "))
                    if isinstance(num, int) and 0 < num <= (int(len(questions))):
                            questions.pop(num-1)
                            answers.pop(num-1)
                            print("Question {} removed".format(num))
                else:
                    raise ValueError
            except ValueError:
                print('\nThat didnt match the format, please try again\n')
    else:
        print("\nInvalid Input, please try again\n")

a1Lim = (utils.rowcol_to_a1(image.get_height(), image.get_width()))
print('heres alim')
a1Lim = ''.join([i for i in a1Lim if not (i in "1234567890")])
print(a1Lim)

pixelSize = int(1200/image.get_width())
print(pixelSize)

set_column_width(sheet2, "A:{}".format(a1Lim), pixelSize)
set_row_height(sheet2, "1:{}".format(image.get_height()), pixelSize)

rows = [*range(0, image.get_height()-1)]
columns = [*range(0, image.get_width()-1)]
print((rows, columns))
#WHERE THE SWAG HAPPENS
rowValues = []
for row in rows:
    print("row: {}".format(row))
    colValues = []
    formValues = []
    for column in columns:
        print("column: {}".format(column))
        color = (image.get_at((column, row)))
        address = utils.rowcol_to_a1(row+1, column+1)
        colValues.append({"userEnteredFormat": {"backgroundColorStyle": {"rgbColor": {"red": color[0] / 255,"green": color[1] / 255,"blue": color[2] / 255}}}})
    rowValues.append({"values": colValues})

print("updating colors")

#formula = '=indirect("Sheet1!B2")<>"hello"'
#requests = {"requests": [{"addConditionalFormatRule": {"index": 0, "rule": {"ranges": [{"sheetId": sheet2.id}],"booleanRule": {"condition": {"type": "CUSTOM_FORMULA","values": [{"userEnteredValue": formula}]},"format": {"backgroundColorStyle": {"rgbColor": {"red": 1, "green": 1, "blue": 1}}}}}}}]}

colorRequest = {"requests": [{"updateCells": {"rows": rowValues,"range": {"sheetId": sheet2.id,"startRowIndex": rows[0],"endRowIndex": rows[-1] + 1,"startColumnIndex": columns[0],"endColumnIndex": columns[-1] + 1},"fields": "*"}}]}
sheet.batch_update(colorRequest)

#PRINTING QUESTIONS
myList = []
for question in questions:
    myList.append([str(question), ''])
while True:
    try:
        startLocation = str(input("\nEnter the cell where you would like the questions to start:"))
        sheet.values_update(
            'Sheet1!{}'.format(startLocation), 
            params={'valueInputOption': 'RAW'}, 
            body={'values': myList}
        )
        print(startLocation)
        rowColShift = utils.a1_to_rowcol(startLocation) #returns tuple, coordinates for cells
        print("rowColShift 1: {}".format(rowColShift))
        rowColShift = (rowColShift[0] - 1, rowColShift[1] + 1) #increases column number for questions, decreases row number as well
        print("rowColShift 2: {}".format(rowColShift))
        rowColShift = utils.rowcol_to_a1(rowColShift[0], rowColShift[1])
        print("rowColShift 3: {}".format(rowColShift))
        columnShift = ''.join([i for i in rowColShift if not (i in "1234567890")]) #returns letter
        rowShift = (int(''.join([i for i in rowColShift if not (i in "ABCDEFGHIJKLMNOPQRSTUVWXYZ")]))) #returns number
        break
    except TypeError:
        print("\nWoops, there was an issue with that cell. Make sure you enter it in A1 notation.\nFor example, if you enter B3, the first question will be on cell B3")

print("\nformatting cells\n")

formatRequests = []
for row in rows:
    for column in columns:
        rowStart = startLocation
        qIndex = random.randint(1, len(questions))
        formatRequests.append(
            {
                "addConditionalFormatRule": {
                    "index": 0,
                    "rule": {
                        "ranges": [
                            {
                                "sheetId": sheet2.id,
                                "startRowIndex": row,
                                "endRowIndex": row + 1,
                                "startColumnIndex": column,
                                "endColumnIndex": column + 1,
                            }
                        ],
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA",            #aIndex: the column (letter) to print, qIndex: the row (number) to print
                                "values": [{"userEnteredValue": '=indirect("Sheet1!{}{}")<>"{}"'.format(columnShift, qIndex + rowShift, answers[qIndex-1])}],
                            },
                            "format": {
                                "backgroundColorStyle": {
                                    "rgbColor": {"red": 1, "green": 1, "blue": 1}
                                }
                            },
                        },
                    },
                }
            }
        )

sheet.batch_update({"requests": formatRequests})

while True:
    try:
        userEmail = input("\nEnter your email to sent this sheet to: ")
        sheet.share(userEmail, 'user', 'writer')
        break
    except gspread.exceptions.APIError:
        print("There was a problem sharing with that email, make sure you typed it correctly and try again")

print('Shared sheet to {}'.format(userEmail))

print("\nProgram has finished, this window will close in 20 seconds\n")

time.sleep(20)
print("Goodbye")

#allow user to pick the cell where the questions start
#save email