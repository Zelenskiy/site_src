#! /usr/bin/env python
# -*- coding: utf-8 -*-

'''
Процес створення API токена
https://www.youtube.com/watch?v=Bf8KHZtcxnA&ab_channel=%D0%94%D0%B8%D0%B4%D0%B6%D0%B8%D1%82%D0%B0%D0%BB%D0%B8%D0%B7%D0%B8%D1%80%D1%83%D0%B9%21
До прав редагувати таблицю додати користувача account@t-sunlight-247017.iam.gserviceaccount.com
(з файла project-fa0cf409504d.json)
 https://habr.com/ru/post/305378/
'''

from pprint import pprint

import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from Spreadsheet import Spreadsheet

# Файл, полученный в Google Developer Console
CREDENTIALS_FILE = 'project-fa0cf409504d.json'
# ID Google Sheets документа (можно взять из его URL)
# spreadsheet_id = '1gIGSxWp-DQ6Cm5KiB-Z76gj4YyN0crjseQQgCetDCtY'

def init():
    # Авторизуемся и получаем service — экземпляр доступа к API
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        CREDENTIALS_FILE,
        ['https://www.googleapis.com/auth/spreadsheets',
         'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)
    return  service

#Шукаємо перший порожній рядок на сторінці missingbook
def searchEmptyRow(idSpreadheet, nameSheet):
    service = init()
    i = 3
    s = service.spreadsheets().values().get(
        spreadsheetId=idSpreadheet,
        range=nameSheet + '!B'+str(3)+':J'+str(10000)+'',
        majorDimension='ROWS'
    ).execute()
    for row in s['values']:
        # print(row)
        i += 1
    return i

def searchOnDate(spreadsheet_id, d1, d2):
    service = init()
    s = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='missingbook!B' + str(3) + ':J' + str(5000) + '',
        majorDimension='ROWS'
    ).execute()
    newList = []
    for row in s['values']:
        if row[1] >= d1 and row[1] <= d2 :
            newList.append(row)
    return newList



def test(spreadsheet_id):
    service = init()
    # read from spreadsheet
    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='workdays!A:E10',
        majorDimension='ROWS'
    ).execute()
    return values

def write(spreadsheet_id, i,text):
    service = init()
    # write to spreadsheet
    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": "missingbook!A"+str(i)+":J"+str(i)+"",
                 "majorDimension": "ROWS",
                 "values": [text]},
                # {"range": "D5:E6",
                #  "majorDimension": "COLUMNS",
                #  "values": [["This is D5", "This is D6"], ["This is E5", "=5+5"]]}
            ]
        }
    ).execute()

def addRow(idSpreadheet, list):
    i = (searchEmptyRow(idSpreadheet, nameSheet="massingbook"))
    write(i, list)

def addBlock(idSpreadheet, nameSheet, lst):
    i = searchEmptyRow(idSpreadheet, nameSheet)
    service = init()
    length = len(lst)
    print(length)
    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=idSpreadheet,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": "missingbook!A" + str(i) + ":J" + str(i+length) + "",
                 "majorDimension": "ROWS",
                 "values": lst},
            ]
        }
    ).execute()

def readBlock(idSpreadheet, nameSheet="massingbook", block='A1:J5000'):
    service = init()
    list = service.spreadsheets().values().get(
        spreadsheetId=idSpreadheet,
        range= nameSheet + '!' + block,
        majorDimension='ROWS'
    ).execute()
    return list

def _read(idSpreadheet, nameSheet):
    service = init()
    i = 3
    s = service.spreadsheets().values().get(
        spreadsheetId=idSpreadheet,
        range=nameSheet + '!A'+str(3)+':N'+str(10000)+'',
        majorDimension='ROWS'
    ).execute()

    return s['values']
def read_settings(idSpreadheet, nameSheet):
    service = init()
    s = service.spreadsheets().values().get(
        spreadsheetId=idSpreadheet,
        range=nameSheet + '!A'+str(1)+':B'+str(10000)+'',
        majorDimension='ROWS'
    ).execute()

    return s['values']

if __name__ == "__main__":
    teachers = ["Труш Ольга Миколаївна", "Зеленський Олександр Станіславович", "Лапко Юлія Миколаївна",
                "Береговець Наталія Василівна", "Богдан Олена Іванівна", "Бойко Лідія Анатоліївна",
                "Бугай Віталій Григорович", "Бушко Валентина Володимирівна", "Герасименко Марія Петрівна",
                "Гнилуша Лілія Володимирівна", "Давиденко Микола Григорович", "Дзюба Людмила Миколаївна",
                "Дитюк Олена Іванівна", "Долиненко Вікторія Андріївна", "Дуденко Олена Юріївна",
                "Звєрєв Василь Юрійович", "Зеленська Галина Олексіївна", "Іванюк Ольга Миколаївна",
                "Карпенко Ганна Вікторівна", "Кашперська Людмила Василівна", "Компанець Любов Григорівна",
                "Коса Тетяна Михайлівна", "Кузуб Галина Федорівна", "Купріян Людмила Володимирівна",
                "Литвякова Наталія Вікторівна", "Лопата Олена Миколаївна", "Мартинова Юлія Сергіївна",
                "Масльонка Володимир Григорович", "Мельник Валентина Михайлівна", "Міна Алла Миколаївна",
                "Мозоль Олена Володимирівна", "Мурза Світлана Миколаївна", "Ніколаєнко Юлія Олександрівна",
                "Опанасенко Ніна Андріївна", "Орсагош Руслан Ілліч", "Пилипенко Лариса Віталіївна",
                "Пилипенко Ольга Павлівна", "Пильник Світлана Володимирівна", "Підвербна Любов Василівна",
                "Плеса Інна Вікторівна", "Пономарчук Валентина Михайлівна", "Ревко Алла Олександрівна",
                "Сало Алла Володимирівна", "Сахно Ярина Дмитрівна", "Собко Валентина Іванівна",
                "Сом Тетяна Миколаївна", "Тарасенко Олена Миколаївна", "Титаренко Альона Сергіївна",
                "Халімон Олександр Анаталійович", "Чава Марія Омелянівна", "Чава Сергій Анатолійович",
                "Шиш Наталія Василівна", "Ведмідь Петро Іванович", "Шемендюк Любов Іванівна",
                "Василенко Ігор Володимирович", "Волошина Марина С.", "Герасименко Віктор Олександрович",
                "Гуляй Людмила Миколаївна", "Гуща Ірина Віталіївна", "Давиденко Ольга Віталіївна",
                "Івашута Наталія Михайлівна", "Йовенко Юлія Вікторівна", "Козачок Наталія Іванівна",
                "Момот Наталія Андріївна", "Пономарчук Валентина  Олексіївна", "Труба Світлана Петрівна",
                "Турченяк Петро Йосипович", ]

    # Читаємо дані про видані свідоцтва в Choippo
    # listAll = _read('1riSyWbebtO2WaY1aqldAbBeHSgy5IMFGi7lVeG0jqkU', "Лист1")  #Свідоцтва
    listAll = _read('1dnvRSD3FlDxaxkDjHrHb-_gRTPTWSUkaampVwawnSos', "Лист1")  #Сертифікати
    for l in listAll:
        if len(l) > 2:
            k = l[2].strip()
            if k != '':
                for t in teachers:
                    if k in t:
                        for a in l:
                            print(a, end="\t")
                        print()

