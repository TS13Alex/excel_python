from py_modul import readlist
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins
from openpyxl.styles import NamedStyle, Border, Side, Alignment, Font
from pathlib import WindowsPath
from pathlib import Path
import argparse



# Создаем комплектовочную ведомость из ведомости покупных формата CSV
# Запускать с парметром :номер ГБНК (468332.124)
# Из файла list.txt берем название устройства для заполнения данных
#DIR_VP = 'C:/Users/tsarev.NIIAEM/Documents/2023/biab200/vp/'
#NAME_VP = 'VP200.csv'
DIR_VP = 'C:/Users/tsarev.NIIAEM/Documents/2023/biab200/vp/'
NAME_VP = 'VP200.csv'
NAME_FILE = 'list.csv'
DOGOVOR = 'Договор: 1526730203022214000241307/2914/21-EП-732/724/610'
IZDELIE = 'Изделие: БИАБ-200ЛИ   ГБНК.566111.024   зав. № 10, 11'
NAME_KOMPL = 'komlpl.xlsx'
DIR_KM = 'C:/Users/tsarev.NIIAEM/Documents/2023/python_project/VPtoKomplekt/'

def cell_style(ws, row, col, al='left', clr='000000', wr=False, sz = 12):  # примменить стиль к ячейке
        bd = Side(style='thin', color=clr)
        cell = ws.cell(row=row, column=col)
        cell.font = Font(name='Times New Roman', size=sz)
        cell.border = Border(left=bd, top=bd, right=bd, bottom=bd)
        cell.alignment = Alignment(
            horizontal=al, vertical='center',  wrap_text=wr)


def shablon(ws, start=1):
    for j in range(1, 5):
            for d, al in [[0, 'center'], [2, 'right'], [4, 'center'], [6, 'left']]:
                cell_style(ws, start+d, j, al, 'ffffff')

    for i, val_sell in [[0, 'Комплектовочный лист сборочной единицы:'], [2, DOGOVOR], [4,IZDELIE]]:
        ws.merge_cells(start_row=start+i, start_column=1, end_row=start+i, end_column=4)
        ws.cell(row=start+i, column=1, value=val_sell)    
    ws.cell(row=start+6, column=1, value=' '+n_dev_name[:40]+'  '+n_dev_gbnk)

    for i in range(start+8, start+40):
        for j in range(2, 5):              
            cell_style(ws, i, j, 'center',sz=14)
        cell_style(ws, i, 1, 'left',sz=14)
                             
        
        
    for i in range(start+40, start+46):
        for j in range(1, 5):              
            cell_style(ws, i, j, 'right', 'ffffff')
        ws.merge_cells(start_row=i, start_column=1, 
                                    end_row=i, end_column=4)
    ws.cell(row=start+41, column=1, value='Ответственный исполнитель _________  _________________  _____________')
    ws.cell(row=start+42, column=1, value='(дата)                (подпись)                     (фамилия)')
    ws.cell(row=start+44, column=1, value='Сотрудник ОМТС _________  _________________  Виноградова Л.Б')
    ws.cell(row=start+45, column=1, value='(дата)                  (подпись)                                   ')

    val_text = ['Наименование ЭРИ ПКИ', 'Кол-во', '№ входного протокола', 'Примечание']
    for j in range (1, 5):
        cell_style(ws, start+7, j, 'center', wr=True)
        ws.cell(row=start+7, column=j, value=val_text[j-1])

if __name__ == "__main__":
    
    parser = argparse.ArgumentParser()
    parser.add_argument('gbnk', nargs='?', default=None)
    args = parser.parse_args()
    gbnk = args.gbnk
    #print(gbnk)
    p = WindowsPath(DIR_VP+NAME_FILE)
    name_dev = readlist.namedev_list(p, gbnk)
    wb = Workbook()
    ws = wb.active
    ws.page_margins = PageMargins(0.5, 0.2, 0.2, 0.2)
    ws.column_dimensions['A'].width = 60
    ws.column_dimensions['B'].width = 7
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 13
    start = 1
    
    for n_dev_gbnk, n_dev_name in name_dev:
        #print(n_dev_gbnk)
        shablon(ws, start)
        cwd = WindowsPath(DIR_VP+NAME_VP)
        p = Path(cwd)
        f = open(p)
        a1 = [[line.split(';')[4], line.split(';')[11]] for line in f if line.split(';')[9] == n_dev_gbnk]
        # ws.page_setup.paperSize = 9
        # ws.page_setup.orientation = 'portrait'
        page = 1 #номер страницы
        d = {}
        for i, j in a1:
            #читаем список и складываем в словарь уникальные имена кол-во суммируем если имя уже есть в словаре
            try:
                num = int(j)*2 #комплектование на 2 стойки
            except ValueError:
                num = 0
            if 'ГБНК' in i:
                 #print(i, j)
                 pass
            else:
                if i in d.keys():
                    d[i] += num
                else:
                    d[i] = num

        all_page = str(len(d)//32 + 1)
        ws.cell(start+6, 3, value='Лист ' + str(page))
        ws.cell(start+6, 4, value='Листов '+ all_page)
        r = 1
        for i in sorted(d):
            if 'Конденсатор' in i:
                 name_elem = 'Кон' + i[11:]
            else:
                 name_elem = i
            row = r + start + 7
            ws.cell(row, 1, value=name_elem)
            ws.cell(row, 2, value=d[i])
            r += 1
            if r > 32:
                 start += 46
                 page += 1
                 ws.cell(start+6, 3, value='Лист ' + str(page))
                 ws.cell(start+6, 4, value='Листов '+ all_page)
                 r = 1
                 shablon(ws, start)
        start += 46
        wb.save(WindowsPath(DIR_KM + NAME_KOMPL))
    
