from sheetutils import *
import datetime

def strdateShortToUniStrdate(d): # 01.09.2020 or 01.09.20 to 2020-09-01
    if len(d) == 8:
        d = d[:6] + '20' + d[6:]
    d = d[6:] + '-' + d[3:5] + '-' + d[:2]
    return d

def psmaker(idSpreadheet, nameSheet="missingbook", d1='2020-09-01', d2='2020-09-30'):
    #
    n = searchEmptyRow(idSpreadheet, nameSheet)
    #
    list = readBlock(idSpreadheet, nameSheet, block='A1:J'+str(n))

    history = []

    for row in list['values']:
        if len(row) < 9:
            row += ["", "", "", "", "", "", ""]
        rec = {}
        dat = strdateShortToUniStrdate(row[1])
        if (dat > d1 or dat == d1) and (dat < d2 or dat == d2):
            rec['D'] = dat
            rec['VT'] = row[2]
            rec['PV'] = row[3]
            rec['P1'] = row[4]
            rec['KL'] = row[5]
            rec['ZT'] = row[6]
            rec['R'] = row[7]
            rec['P2'] = row[8]
            history += [rec]


    '''
    s = str(sheet['A' + str(n)].value)
    history = []
    while s is not None:
        rec = {}
        n += 1
        s = str(sheet['A' + str(n)].value)
        dat = (sheet['B' + str(n)].value)
        # dat = datetime.datetime.strptime(datS,'%Y-%m-%d %h:%m:%s')
        # d1 = datetime.datetime.strptime('01.09.2019', '%d.%m.%Y')
        # d2 = datetime.datetime.strptime('30.09.2019', '%d.%m.%Y')

        try:
            k = int(s)
        except:
            break
        if (dat > d1 or dat == d1) and (dat < d2 or dat == d2):
            rec['D'] = dat

            rec['VT'] = str(sheet['C' + str(n)].value)
            rec['PV'] = str(sheet['D' + str(n)].value)
            rec['P1'] = str(sheet['E' + str(n)].value)
            rec['KL'] = str(sheet['F' + str(n)].value)
            rec['ZT'] = str(sheet['G' + str(n)].value)
            rec['P2'] = str(sheet['I' + str(n)].value)
            tmp = str(sheet['J' + str(n)].value)
            if tmp == 'pk':
                rec['poch_kl'] = True
            else:
                rec['poch_kl'] = False
            tmp = str(sheet['L' + str(n)].value)
            if tmp == 'kl_ker':
                rec['kl_ker'] = True
            else:
                rec['kl_ker'] = False
            history += [rec]
    print(n)

'''

    teachers = []
    for h in history:
        # Визначаємо кількість вчителів
        title = h['VT'] + '~' + h['PV']
        reason = h['PV']
        # rec = {}
        if not title in teachers:
            teach = h['VT']
            subjects = []
            zamteachers = []
            rec = {}
            for h2 in history:
                title2 = h2['VT'] + '~' + h2['PV']
                rec4 = {}
                if title == title2:
                    subject = h2['P1']
                    if not subject in subjects:
                        klases = []
                        for h3 in history:
                            title3 = h3['VT'] + '~' + h3['PV']
                            if title2 == title3 and h3['P1'] == h2['P1']:
                                klas = h3['KL']
                                if not klas in klases:
                                    klases += [klas]
                        rec1 = {}
                        rec1['subject'] = subject
                        rec1['classes'] = klases
                        if not rec1 in subjects:
                            subjects += [rec1]

            rec['title'] = title
            rec['reason'] = reason
            rec['subjects'] = subjects
            rec['teach'] = teach
            # rec['zamteach'] = [rec3]
            if not rec in teachers:
                teachers += [rec]

    for t in teachers:
        count = 0
        date = []
        dss = []
        kl_ker = False
        poch_kl = False
        for h in history:
            title = h['VT'] + '~' + h['PV']

            if title == t['title']:
                count += 1
                ds = date_to_dd_mm(datetime.datetime.strptime(h['D'], '%Y-%m-%d'))
                if not ds in dss:
                    dss += [ds]
                date += [ds]
                # if h['kl_ker']:
                #     kl_ker = True
                # if h['poch_kl']:
                #     poch_kl = True

        # t['kl_ker'] = (kl_ker)
        # t['poch_kl'] = (poch_kl)
        cc, zzz = teach_to_zteach(history, t)
        t['zteachers'] = {'countZteach': cc, 'zts': zzz}

        t['count'] = count
        t['min'] = min_list(date)
        t['max'] = max_list(date)
        t['days'] = len(dss)

    listExcel = []

    for i, t in enumerate(teachers):
        rec = {}
        rec['N'] = str(i + 1)
        rec['teach'] = t['teach']
        subjects = t['subjects']
        s = ''
        for subject in subjects:
            klss = ''
            for kl in subject['classes']:
                klss += kl
            klss = klss.replace('-', '')
            s += subject['subject'] + '(' + klss + '),'
        rec['subject'] = s[:-1]
        rec['dmin'] = t['min']
        rec['dmax'] = t['max']
        rec['days'] = t['days']
        rec['count'] = t['count']
        rec['reason'] = t['reason']
        rec['zteachers'] = t['zteachers']
        # rec['kl_ker'] = t['kl_ker']
        # rec['poch_kl'] = t['poch_kl']

        listExcel += [rec]
    n = 7
    sheet = [["№ п/п", "Відсутній вчитель", "Предмет, клас", "З якого дня відсутній", "По який день відсутній",
                           "Пропущен. робочих днів", "Пропущених годин", "Вчитель, який заміняє", "Предмет, клас",
                           "Прочитан. годин", "Примітка"]]
    for t in listExcel:
        row = ['','','','','','','','','','','']
        row[0] = t['N']
        row[1] = t['teach']
        row[2] = t['subject']
        row[3] = t['dmin']
        row[4] = t['dmax']
        row[5] = t['days']
        row[6] = t['count']
        row[7] = ''
        row[8] = ''
        row[9] = ''
        row[10] = t['reason']

        zts = t['zteachers']['zts']
        for zt in zts:
            z = zt['zteach']
            row[7] = z
            countLesson = zt['countLesson']
            row[9] = countLesson
            s = ''
            for subj in zt['subjects']:
                c = ''
                for cl in subj['classes']:
                    c += cl
                c = c.replace('-', '')
                s = subj['subject'] + '(' + c + ')'
                ss = s + '/' + str(subj['countKlas']) + 'год/'
                if ss[:3].upper() == 'ІНД':
                    ss += '+20%'
                row[8] = ss
                n += 1
                sheet.append(row)
                row = ['', '', '', '', '', '', '', '', '', '', '']
            # sheet.append(row)
            # row = ['', '', '', '', '', '', '', '', '', '', '']

        n += 1
        sheet.append(row)
        row = ['', '', '', '', '', '', '', '', '', '', '']

    # Здійснюємо переноси при занадто довгому рядку у стовпці D (2)
    # кількість символів, що влазить у стовпецю 25
    for row in range(1, len(sheet)):
        if sheet[row][1] != "":
            s = sheet[row][2]
            if len(s) + 20:
                ss = s.split(',')
                for i, s0 in enumerate(ss):
                    sheet[row + i][2] = s0




    return sheet

def date_to_dd_mm(d):
    day = str(d.day)
    if len(day) == 1:
        day = '0' + day
    month = str(d.month)
    if len(month) == 1:
        month = '0' + month
    return day + '.' + month + '.'

def min_list(ls):
    min = ls[0]
    for l in ls:
        if l < min:
            min = l
    return min


def max_list(ls):
    max = ls[0]
    for l in ls:
        if l > max:
            max = l
    return max

def teach_to_zteach(history, t):
    count = 0
    zteachers = []
    for h in history:
        title = t['title']
        htt = title.split('~')
        vt = htt[0]
        pv = htt[1]
        if h['VT'] == vt and h['PV'] == pv:
            zt = h['ZT']
            count += 1
            countSubj, sses = zteach_to_subjects(history, title, zt)
            if not sses in zteachers:
                zteachers += [sses]

    count = len(zteachers)
    return count, zteachers

def zteach_to_subjects(history, title, zt):
    subjects = {'zteach': zt}
    subj = []
    count = 0
    for h in history:
        htt = title.split('~')
        vt = htt[0]
        pv = htt[1]
        if h['VT'] == vt and h['PV'] == pv and h['ZT'] == zt:
            count += 1
            if not h['P1'] in subj:
                subj += [h['P1']]
    subjects['countLesson'] = count
    pr = []
    for s in subj:
        dic = {}
        if not s in pr:
            c, clss = subject_to_classes(history, title, zt, s)
            dic['subject'] = s
            dic['classes'] = clss
            dic['countKlas'] = c
            pr += [dic]
    subjects['subjects'] = pr
    return count, subjects

def subject_to_classes(history, title, zt, sbs):
    classes = []
    count = 0
    for h in history:
        htt = title.split('~')
        vt = htt[0]
        pv = htt[1]
        if h['VT'] == vt and h['PV'] == pv and h['ZT'] == zt and h['P1'] == sbs:
            count += 1
            if not h['KL'] in classes:
                classes += [h['KL']]
    return count, classes

if __name__ == "__main__":
    pass
