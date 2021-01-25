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
spreadsheet_id = '1gIGSxWp-DQ6Cm5KiB-Z76gj4YyN0crjseQQgCetDCtY'

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
        range=nameSheet + '!B'+str(3)+':J'+str(5000)+'',
        majorDimension='ROWS'
    ).execute()
    for row in s['values']:
        # print(row)
        i += 1
    return i

def searchOnDate(d1, d2):
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



def test():
    service = init()
    # read from spreadsheet
    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='workdays!A:E10',
        majorDimension='ROWS'
    ).execute()
    return values

def write(i,text):
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
        spreadsheetId=spreadsheet_id,
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
    return list;

if __name__ == "__main__":
    # addRow(['text2', '1', '2', '3', '4', '5', '6', '7', '8', '9'])
    # addBlock([
    #     ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'],
    #     ['00', '10', '20', '30', '40', '50', '60', '70', '80'],
    #     ['000', '100', '200', '300', '400', '500', '600', '700', '800', '900'],])
    print(readBlock(spreadsheet_id, "missingbook", "A879:J881")['values'])
    # addSheet()

