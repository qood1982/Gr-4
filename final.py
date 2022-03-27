import csv
import re
from majed_modules.multi import mysort
from poppler import load_from_file
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.formatting import formatting
from majed_modules.multi import basic_salary_dictionary


def mysort(item):
    date_regex = re.compile(r"Payroll for (\d*)/ (\d{4})")
    date = date_regex.search(item)
    m = date.group(1)
    y = date.group(2)
    m_y = y[2:4] + m
    return int(m_y)


paths = Path('./Assets/Majed Ghanim')
list_of_all_pages = []
with open(f'./final/out/{paths.name}.txt', 'w') as file:

    for p in paths.iterdir():
        if p.suffix == '.pdf':
            doc = load_from_file(p)
        if doc.pages == 2:
            dual0 = doc.create_page(0)
            dual1 = doc.create_page(1)
            txt1 = dual0.text()
            txt2 = dual1.text()
            mix = txt1 + txt2
            t = file.write(mix)
            list_of_all_pages.append(mix)
        else:
            singel = doc.create_page(0)
            txt = singel.text()
            t = file.write(txt)

            list_of_all_pages.append(txt)

list_of_all_pages.sort(key=mysort)


def basic_salary_dictionary(args):
    Basic_Salary_Dictionary = {}
    for bs in args:
        # bsx = re.search(r'^0001[\s]*(.*?)[\s]*([\d,.]*)[\s]*(\d{6})', bs,
        # re.MULTILINE)
        for l in bs.splitlines():
            if re.search(r'^0001', l):
                s = l.split()
                bs_amount = s[-2]
                bs_amount = float(bs_amount.replace(',', ''))
                bs_date = s[-1]
                if bs_date not in Basic_Salary_Dictionary:
                    Basic_Salary_Dictionary[bs_date] = round(bs_amount, 3)

    return Basic_Salary_Dictionary


bsd = basic_salary_dictionary(list_of_all_pages)

total = 0
count = 0
crossover = {}
with open(f'./out/{paths.name}.csv', 'w') as cv:
    w = csv.writer(cv)
    w.writerow([u'التاريخ', u'نوع البدل', u'الحسميات', u'مرجع'])
    for page in list_of_all_pages:
        payroll = re.search(r'Payroll for (\d*/ \d{4})', page)
        annual = re.search(r'^3000.+', page, re.MULTILINE)
        name = re.search(r'Name : .*', page)
        modefied_payroll = payroll.group(1).replace(('/'), '')
        modefied_payroll = modefied_payroll.replace((' '), '')
        m = modefied_payroll[0:2]
        y = modefied_payroll[2:]
        m_y = y + m
        count = 0
        aux = []
        for line in page.splitlines():
            s = line.split()
            if re.search(r'^1110|^1111|^1115|^1113|^1320', line):

                if re.search(r'^1110.+-', line):
                    date10 = s[-1]
                    amount10 = float(s[-2].replace(',', ''))
                    disc10 = u'بدل وردية متغيرة ١٠٪'
                    total += amount10
                    w.writerow([date10, disc10, amount10, payroll.group(1)])
                    print(date10, disc10, amount10, payroll.group(1))

                elif re.search(r'^1115.+-', line):
                    date15 = s[-1]
                    amount15 = float(s[-2].replace(',', ''))
                    disc15 = u'تعويض جدول عمل ٥٪'
                    total += amount15
                    print(date15, disc15, amount15, payroll.group(1))
                    w.writerow([date15, disc15, amount15, payroll.group(1)])

                elif re.search(r'^1111', line):
                    if annual:
                        if s[-1] in m_y:  #Same Month PayRoll
                            date = s[-1]
                            amount = float(s[-2].replace(',', ''))
                            amount = round(amount - bsd.get(date) * 0.1, 3)
                            disc = u'بدل وردية متغيرة ١٠٪'
                            w.writerow([date, disc, amount, payroll.group(1)])
                            total += amount
                            if date not in crossover:
                                crossover[date] = amount
                            print(date, disc, amount)

                    elif s[-1] not in m_y:  # Not Same Month PayRoll
                        date = s[-1]
                        amount = float(s[-2].replace(',', ''))
                        aux.append(amount)
                        disc = u'بدل وردية متغيرة ١٠٪'
                        if len(aux) == 2:
                            amount = round(aux[0] + aux[1], 3)

                            if amount > 0:
                                amount = round(
                                    crossover.get(date, 0) + amount, 3)
                                w.writerow([date, disc, amount])
                                total += amount
                                print(date, disc, amount)

                            else:
                                w.writerow(
                                    [date, disc, amount,
                                     payroll.group(1)])
                                total += amount
                                print(date, disc, amount)

                        elif len(aux) == 4:
                            amount = round(aux[0] + aux[1], 3)
                            if amount > 0:

                                amount = round(
                                    crossover.get(date, 0) + amount, 3)
                                w.writerow(
                                    [date, disc, amount,
                                     payroll.group(1)])
                                total += amount
                                print(date, disc, amount)
                            else:
                                w.writerow(
                                    [date, disc, amount,
                                     payroll.group(1)])
                                total += amount
                                print(date, disc, amount)

                        elif len(aux) == 6:
                            amount = round(aux[0] + aux[1], 3)
                            if amount > 0:

                                amount = round(
                                    crossover.get(date, 0) + amount, 3)
                                w.writerow(
                                    [date, disc, amount,
                                     payroll.group(1)])
                                total += amount
                                print(date, disc, amount)
                            else:
                                w.writerow(
                                    [date, disc, amount,
                                     payroll.group(1)])
                                total += amount
                                print(date, disc, amount)

                        elif len(aux) == 8:
                            amount = round(aux[-2] + aux[-1], 3)
                            if amount > 0:

                                amount = round(
                                    crossover.get(date, 0) + amount, 3)
                                w.writerow(
                                    [date, disc, amount,
                                     payroll.group(1)])
                                total += amount
                                print(date, disc, amount)
                            else:
                                w.writerow(
                                    [date, disc, amount,
                                     payroll.group(1)])
                                total += amount
                                print(date, disc, amount)

                elif re.search(r'^1113.+-', line):
                    date = s[-1]
                    amount = float(s[-2].replace(',', ''))
                    disc = u'تعويض جدول عمل ٥٪'
                    total += amount
                    w.writerow([date, disc, amount, payroll.group(1)])
                    print(date, disc, amount)
                elif re.search(r'^1113', line):
                    if annual:
                        if s[-1] in m_y:  #Same Month PayRoll
                            date = s[-1]
                            amount = float(s[-2].replace(',', ''))
                            amount = round(amount - bsd.get(date) * 0.05, 3)
                            disc = u'تعويض جدول عمل ٥٪'
                            total += amount
                            w.writerow([date, disc, amount, payroll.group(1)])

                elif re.search(r'^1320', line):
                    if '-' in line:
                        date = s[-1]
                        amount = float(s[-2].replace(',', ''))
                        disc = u'بدل طبيعة عمل'
                        total += amount
                        w.writerow([date, disc, amount, payroll.group(1)])
                    elif annual:
                        if s[-1] in m_y:  #Same Month PayRoll
                            date = s[-1]
                            amount = float(s[-2].replace(',', ''))
                            amount = round(amount - bsd.get(date) * 0.2, 3)
                            disc = u'بدل طبيعة عمل'
                            total += amount
                            w.writerow([date, disc, amount, payroll.group(1)])

                        # if aux:
                        #     print(aux)
    w.writerow([
        None,
        '',
        f'{round(total, 3)}',
        ': المجموع',
    ])
