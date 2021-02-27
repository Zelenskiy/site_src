#!/usr/bin/env python3

# Author: Ioann Volkov (volkov.ioann@gmail.com)
# This module uses Google Sheets API v4 (and Google Drive API v3 for sharing spreadsheets)

# (!) Disclaimer
# This is NOT a full-functional wrapper over Sheets API v4.
# This module was created just for https://telegram.me/TimeManagementBot and habrahabr article

from pprint import pprint
import httplib2
import apiclient.discovery
import googleapiclient.errors
from oauth2client.service_account import ServiceAccountCredentials

from Spreadsheet import Spreadsheet
from pz import psmaker
from sheetutils import read_settings


def htmlColorToJSON(htmlColor):
    if htmlColor.startswith("#"):
        htmlColor = htmlColor[1:]
    return {"red": int(htmlColor[0:2], 16) / 255.0, "green": int(htmlColor[2:4], 16) / 255.0,
            "blue": int(htmlColor[4:6], 16) / 255.0}


# === Tests for class Spreadsheet ===

GOOGLE_CREDENTIALS_FILE = 'project-fa0cf409504d.json'


def testCreateSpreadsheet():
    ss = Spreadsheet(GOOGLE_CREDENTIALS_FILE, debugMode=True)
    ss.create("Preved medved", "Тестовый лист")
    ss.shareWithEmailForWriting("zelensk1973@gmail.com")


def testSetSpreadsheet():
    ss = Spreadsheet(GOOGLE_CREDENTIALS_FILE, debugMode=True)
    ss.setSpreadsheetById('19SPK--efwYq9pZ7TvBYtFItxE0gY3zpfR5NykOJ6o7I')
    print(ss.sheetId)


def testAddSheet():
    ss = Spreadsheet(GOOGLE_CREDENTIALS_FILE, debugMode=True)
    ss.setSpreadsheetById('19SPK--efwYq9pZ7TvBYtFItxE0gY3zpfR5NykOJ6o7I')
    try:
        print(ss.addSheet("Я лолка №1", 500, 11))
    except googleapiclient.errors.HttpError:
        print("Could not add sheet! Maybe sheet with same name already exists!")


def testSetDimensions():
    ss = Spreadsheet(GOOGLE_CREDENTIALS_FILE, debugMode=True)
    ss.setSpreadsheetById('19SPK--efwYq9pZ7TvBYtFItxE0gY3zpfR5NykOJ6o7I')
    ss.prepare_setColumnWidth(0, 500)
    ss.prepare_setColumnWidth(1, 100)
    ss.prepare_setColumnsWidth(2, 4, 150)
    ss.prepare_setRowHeight(6, 230)
    ss.runPrepared()


def testGridRangeForStr():
    ss = Spreadsheet(GOOGLE_CREDENTIALS_FILE, debugMode=True)
    ss.setSpreadsheetById('19SPK--efwYq9pZ7TvBYtFItxE0gY3zpfR5NykOJ6o7I')
    res = [ss.toGridRange("A3:B4"),
           ss.toGridRange("A5:B"),
           ss.toGridRange("A:B")]
    correctRes = [
        {"sheetId": ss.sheetId, "startRowIndex": 2, "endRowIndex": 4, "startColumnIndex": 0, "endColumnIndex": 2},
        {"sheetId": ss.sheetId, "startRowIndex": 4, "startColumnIndex": 0, "endColumnIndex": 2},
        {"sheetId": ss.sheetId, "startColumnIndex": 0, "endColumnIndex": 2}]
    print("GOOD" if res == correctRes else "BAD", res)


def testSetCellsFormat():
    ss = Spreadsheet(GOOGLE_CREDENTIALS_FILE, debugMode=True)
    ss.setSpreadsheetById('19SPK--efwYq9pZ7TvBYtFItxE0gY3zpfR5NykOJ6o7I')
    ss.prepare_setCellsFormat("B2:E7", {"textFormat": {"bold": True}, "horizontalAlignment": "CENTER"})
    ss.runPrepared()


def testPureBlackBorder():
    ss = Spreadsheet(GOOGLE_CREDENTIALS_FILE, debugMode=True)
    ss.setSpreadsheetById('19SPK--efwYq9pZ7TvBYtFItxE0gY3zpfR5NykOJ6o7I')
    ss.requests.append({"updateBorders": {
        "range": {"sheetId": ss.sheetId, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0,
                  "endColumnIndex": 3},
        "bottom": {"style": "SOLID", "width": 3, "color": {"red": 0, "green": 0, "blue": 0}}}})
    ss.requests.append({"updateBorders": {
        "range": {"sheetId": ss.sheetId, "startRowIndex": 2, "endRowIndex": 3, "startColumnIndex": 0,
                  "endColumnIndex": 3},
        "bottom": {"style": "SOLID", "width": 3, "color": {"red": 0, "green": 0, "blue": 0, "alpha": 1.0}}}})
    ss.requests.append({"updateBorders": {
        "range": {"sheetId": ss.sheetId, "startRowIndex": 3, "endRowIndex": 4, "startColumnIndex": 1,
                  "endColumnIndex": 4},
        "bottom": {"style": "SOLID", "width": 3, "color": {"red": 0, "green": 0, "blue": 0.001}}}})
    ss.requests.append({"updateBorders": {
        "range": {"sheetId": ss.sheetId, "startRowIndex": 4, "endRowIndex": 5, "startColumnIndex": 2,
                  "endColumnIndex": 5},
        "bottom": {"style": "SOLID", "width": 3, "color": {"red": 0.001, "green": 0, "blue": 0}}}})
    ss.runPrepared()
    # Reported: https://code.google.com/a/google.com/p/apps-api-issues/issues/detail?id=4696


def testUpdateCellsFieldsArg():
    ss = Spreadsheet(GOOGLE_CREDENTIALS_FILE, debugMode=True)
    ss.setSpreadsheetById('19SPK--efwYq9pZ7TvBYtFItxE0gY3zpfR5NykOJ6o7I')
    ss.prepare_setCellsFormat("B2:B2", {"textFormat": {"bold": True}, "horizontalAlignment": "CENTER"},
                              fields="userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment")
    ss.prepare_setCellsFormat("B2:B2", {"backgroundColor": htmlColorToJSON("#00CC00")},
                              fields="userEnteredFormat.backgroundColor")
    ss.prepare_setCellsFormats("C4:C4", [[{"textFormat": {"bold": True}, "horizontalAlignment": "CENTER"}]],
                               fields="userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment")
    ss.prepare_setCellsFormats("C4:C4", [[{"backgroundColor": htmlColorToJSON("#00CC00")}]],
                               fields="userEnteredFormat.backgroundColor")
    pprint(ss.requests)
    ss.runPrepared()
    # Reported: https://code.google.com/a/google.com/p/apps-api-issues/issues/detail?id=4697

def tmp_test():
    settings = read_settings(idSpreadheet="1gIGSxWp-DQ6Cm5KiB-Z76gj4YyN0crjseQQgCetDCtY", nameSheet="settings")
    for r in settings:
        if len(r)>1:
            if r[0] == "Коротка назва закладу освіти в називному відмінку":
                name_school = r[1]
            if r[0] == "Посада відповідального за ПЗ":
                posada = r[1]
            if r[0] == "Ініціали, прізвище відповідального за ПЗ":
                piniciali = r[1]
    print(name_school)
    print(posada)
    print(piniciali)


# This function creates a spreadsheet as https://telegram.me/TimeManagementBot can create, but with manually specified data
def pzCreate(idSpreadheet, nameSheet="missingbook", d1='2020-09-01', d2='2020-09-30'):
    # Назва організації
    name_school = "Куликівський ЗЗСО І-ІІІ ст."
    posada = "Заступник директора з НВР"
    piniciali = "Зеленський О.С."
    # https://docs.google.com/spreadsheets/d/1gIGSxWp-DQ6Cm5KiB-Z76gj4YyN0crjseQQgCetDCtY/edit
    settings = read_settings(idSpreadheet, "settings")

    for r in settings:
        if len(r)>1:
            if r[0] == "Коротка назва закладу освіти в називному відмінку":
                name_school = r[1]
            if r[0] == "Посада відповідального за ПЗ":
                posada = r[1]
            if r[0] == "Ініціали, прізвище відповідального за ПЗ":
                piniciali = r[1]



    docTitle = "ПЗ"
    sheetTitle = "PZ"
    title = "ПОЯСНЮВАЛЬНА ЗАПИСКА"
    subTitle_1 = "по виплаті заробітної плати вчителям за заміну відсутніх учителів"
    subTitle_2 = "(по хворобі чи інших обставинах) за період з " + d1[8:10] + '.' + d1[5:7] + '.' + d1[0:4] + " по " + d2[8:10] + '.' + d2[5:7] + '.' + d2[0:4] + ""
    values = psmaker(idSpreadheet, nameSheet=nameSheet, d1=d1, d2=d2)
    rowCount = len(values) - 1
    ss = Spreadsheet(GOOGLE_CREDENTIALS_FILE, debugMode=True)
    ss.create(docTitle, sheetTitle, rows=rowCount + 10, cols=12, locale="ru_RU", timeZone="Europe/Moscow")
    ss.shareWithAnybodyForWriting()

    ss.prepare_setColumnWidth(0, 400)
    ss.prepare_setColumnWidth(1, 200)
    ss.prepare_setColumnsWidth(2, 3, 165)
    ss.prepare_setColumnWidth(4, 100)
    # ss.prepare_mergeCells("A1:E1")  # Merge A1:E1



    ss.prepare_setValues("B1:B1", [[name_school]])
    ss.prepare_setValues("B2:B2", [[title]])
    ss.prepare_setValues("B3:B3", [[subTitle_1]])
    ss.prepare_setValues("B4:B4", [[subTitle_2]])
    ss.prepare_setValues("B6:L%d" % (rowCount + 100), values)

    ss.prepare_setColumnWidth(0, 5)
    ss.prepare_setColumnWidth(1, 35)
    ss.prepare_setColumnWidth(2, 105)
    ss.prepare_setColumnWidth(3, 225)
    ss.prepare_setColumnWidth(4, 45)
    ss.prepare_setColumnWidth(5, 45)
    ss.prepare_setColumnWidth(6, 40)
    ss.prepare_setColumnWidth(7, 40)
    ss.prepare_setColumnWidth(8, 140)
    ss.prepare_setColumnWidth(9, 205)
    ss.prepare_setColumnWidth(10, 35)
    ss.prepare_setColumnWidth(11, 95)

    ss.prepare_setRowHeight(0, 20)
    ss.prepare_setRowHeight(1, 30)
    ss.prepare_setRowHeight(2, 20)
    ss.prepare_setRowHeight(3, 20)
    ss.prepare_setRowHeight(4, 5)
    ss.prepare_setRowHeight(5, 85)
    ss.prepare_setRowHeight(6, 20)
    ss.prepare_setRowHeight(7, 20)

    ss.prepare_mergeCells("B2:L2")
    ss.prepare_mergeCells("B3:L3")
    ss.prepare_mergeCells("B4:L4")

    # rowColors2 = [colorsForCategories[valueRow[0]] for valueRow in values2[1:]]

    ss.prepare_setCellsFormat("B2:B2", {"textFormat": {"fontSize": 18, "bold": True},
                                        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"})
    ss.prepare_setCellsFormat("B3:B3", {"textFormat": {"fontSize": 12, "bold": True},
                                        "horizontalAlignment": "CENTER"})
    ss.prepare_setCellsFormat("B4:B4", {"textFormat": {"fontSize": 12, "bold": True},
                                        "horizontalAlignment": "CENTER"})


    ss.prepare_setCellsFormat("E6:H6", {"textFormat": {"fontSize": 10, "bold": True},
                                        "horizontalAlignment": "CENTER", "textRotation": {"angle": 90}})
    ss.prepare_setCellsFormat("K6:L6", {"textFormat": {"fontSize": 10, "bold": True},
                                        "horizontalAlignment": "CENTER", "textRotation": {"angle": 90},

                                        })

    ss.prepare_setCellsFormat("B6:D6", {"textFormat": {"fontSize": 10, "bold": True},
                                        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                                        "wrapStrategy": "LEGACY_WRAP"})
    ss.prepare_setCellsFormat("E6:H6", {"textFormat": {"fontSize": 9, "bold": True},
                                        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                                        "wrapStrategy": "LEGACY_WRAP", "textRotation": {"angle": 90}})
    ss.prepare_setCellsFormat("I6:J6", {"textFormat": {"fontSize": 10, "bold": True},
                                        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                                        "wrapStrategy": "LEGACY_WRAP"})


    ss.prepare_setCellsFormat("K6:K6", {"textFormat": {"fontSize": 10, "bold": True},
                                        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                                        "wrapStrategy": "LEGACY_WRAP", "textRotation": {"angle": 90}})
    ss.prepare_setCellsFormat("L6:L6", {"textFormat": {"fontSize": 10, "bold": True},
                                        "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                                        "wrapStrategy": "LEGACY_WRAP"})

    ss.requests.append({"updateBorders": {
        "range": {"sheetId": ss.sheetId, "startRowIndex": 5, "endRowIndex": rowCount + 7, "startColumnIndex": 1,
                  "endColumnIndex": 12},
        "bottom": {"style": "SOLID", "width": 1, "color": htmlColorToJSON("#000001")},
        "top": {"style": "SOLID", "width": 1, "color": htmlColorToJSON("#000001")},
        "left": {"style": "SOLID", "width": 1, "color": htmlColorToJSON("#000001")},
        "right": {"style": "SOLID", "width": 1, "color": htmlColorToJSON("#000001")},
        "innerHorizontal": {"style": "SOLID", "width": 1, "color": htmlColorToJSON("#000001")},
        "innerVertical": {"style": "SOLID", "width": 1, "color": htmlColorToJSON("#000001")},

    }})

    ss.prepare_setValues("C"+str(rowCount + 9)+":C" + str(rowCount + 9) + "", [[posada]])
    ss.prepare_setValues("J"+str(rowCount + 9)+":J" + str(rowCount + 9) + "", [[piniciali]])

    # ss.prepare_setValues("G1:G1", [["Категории"]])
    # ss.prepare_setValues("B6:L%d" % (rowCount2 + 3), values2)

    ss.runPrepared()

    print(ss.getSheetURL())
    return (ss.getSheetURL())


if __name__ == "__main__":
    tmp_test()
