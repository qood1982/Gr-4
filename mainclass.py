from poppler import load_from_file
from pathlib import Path
import re


class PayRoll:
    def __init__(self, dirs):
        self.dirs = dirs
        self.basic_salary_dictionary()
        self.list_all_pages()

    def __mysort(self, item):
        date_regex = re.compile(r"Payroll for (\d*)/ (\d{4})")
        date = date_regex.search(item)
        m = date.group(1)
        y = date.group(2)
        m_y = y[2:4] + m
        return int(m_y)

    def list_all_pages(self):
        paths = Path(self.dirs)
        owner = paths.name
        list_of_all_pages = []
        for p in paths.iterdir():
            if p.suffix == '.pdf':
                doc = load_from_file(p)
            if doc.pages == 2:
                dual0 = doc.create_page(0)
                dual1 = doc.create_page(1)
                txt1 = dual0.text()
                txt2 = dual1.text()
                mix = txt1 + txt2
                list_of_all_pages.append(mix)
            else:
                singel = doc.create_page(0)
                txt = singel.text()
                list_of_all_pages.append(txt)

        list_of_all_pages.sort(key=self.__mysort)
        return (list_of_all_pages, owner)

    def basic_salary_dictionary(self):
        Basic_Salary_Dictionary = {}
        for bs in self.dirs:
            # bsx = re.search(r'^0001[\s]*(.*?)[\s]*([\d,.]*)[\s]*(\d{6})', bs,
            # re.MULTILINE)
            for l in bs.splitlines():
                if re.search(r'^0001', l):
                    s = l.split()
                    bs_amount = s[-2]
                    bs_amount = float(bs_amount.replace(',', ''))
                    bs_date = s[-1]
                    if bs_date not in Basic_Salary_Dictionary:
                        Basic_Salary_Dictionary[bs_date] = bs_amount

        return Basic_Salary_Dictionary


majed = PayRoll('./Assets/Majed Ghanim')
