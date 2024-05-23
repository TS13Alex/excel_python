'''генерирует список для заполнения в'''
#from pathlib import WindowsPath
from csv import reader
from json import load
from pathlib import WindowsPath

def add_empty_row(s:list):
    '''add empty row if last row of list general name'''
    k1 = 23
    k2 = 25
    n_st = []
    for i, j in enumerate (s):
        if (i+1 - k1) % k2 == 0 and j[0] == '' and j[1] != '':
            n_st.append(['', '', '', ''])
        n_st.append(j)
    if n_st[-1] == ['', '', '', '']: #if last element == empty, del them
        del n_st[-1]
    while (len(n_st) - k1) % k2 !=0:
        n_st.append(['', '', '', ''])
    return n_st

def list_new_pe (list_el:list, list_name:dict):
    '''generate data for list of elements PE'''
    empty_line = ['', '', '', '']
    list_el.append(empty_line)
    a3 = []
    a_new = []
    f_name ='str'
    for head in list_el:
        if f_name not in head[1]:
            len_a3 = len(a3)
            match len_a3:
                case 0:
                    pass
                case 1:
                    a_new.extend(a3)
                case _:
                    if a_new[-1:] != [empty_line]:
                        a_new.append(empty_line)
                    a_new.append(['', list_name[f_name][1], '', ''])
                    len_name = len(f_name) + 1
                    for elem in a3:
                        elem[1] = elem[1][len_name:]
                        a_new.append(elem)
                    a_new.append(empty_line)
            a3 = []
            for name in list_name:
                if name in head[1]:
                    f_name = name
                    break
            else:
                f_name ='str'
        a3.append(head)
    return a_new

def listing_pe (n_gbnk, vp):
    '''возвращает словарь значений ПЭ из CSV файла'''
    with open (vp, encoding='cp1251', newline='') as vpcsvfile:
        rowreader = reader(vpcsvfile, delimiter=';')
        list_elem = []
        for row in rowreader:
            data_pe = [row[3], row[4], row[10], row[12]]
            if row[3] != '' and row[9] == n_gbnk and data_pe not in list_elem:
                list_elem.append(data_pe)
    return list_elem

def generate_data_pe (vp, gbnk):
    '''read vp file and make data for pe3 doc'''
    m = WindowsPath(__file__)
    t_name = vp.stem
    t = vp.with_name(t_name[3:] + ".json")
    #t = vp[3:]+".json"
    with open(t, "r", encoding="utf-8") as fh:
        dev_json = load(fh) # download gbnk from file
        #print(dev_json)
    key = gbnk
    if key in dev_json:
        print (dev_json[key])
    t = m.with_name("name.json")
    with open(t, "r", encoding="utf-8") as fh:
        name_json = load(fh) # download gbnk from file
    data_dev = listing_pe(key, vp)
    new_data_dev = list_new_pe(data_dev, name_json)
    finell_data_dev = add_empty_row(new_data_dev)
    return finell_data_dev

#=====for BOM

def read_bom_csv(name_file: str):
    '''read csv file return list row'''
    with name_file.open (encoding='cp1251', newline='') as vpcsvfile:
        rowreader = reader(vpcsvfile, delimiter=',')
        next(rowreader) # delete heading string
        next(rowreader)
        readed_list = []
        for row in rowreader:
            data_pe = [row[0], row[1], row[2], row[3]]
            readed_list.append(data_pe)
        name_and_family = [row[4], row[5], row[6], row[7], row[8], row[9], row[10]]
    return (readed_list, name_and_family)



def union_same_element(list_elements: list):
    '''union same element'''
    new_list_elem = []
    head = list_elements[0]
    for row in list_elements[1:]:
        if head[1] == row[1] and head[3] == row[3]:
            head[0] += ', ' + row[0]
        else:
            new_list_elem.append(head)
            head = row
    new_list_elem.append(head)
    return new_list_elem

def change_quantity(list_elements: list):
    '''count number element'''
    pozition_number = list_elements[0].split(', ')
    quantity = len(pozition_number)
    match quantity:
        case 0:
            pass
        case 1:
            pass
        case 2:
            list_elements[2] = str(quantity)
        case _:
            list_elements[2] = str(quantity)
            list_elements[0] = pozition_number[0] + ' - ' + pozition_number[-1]
    return list_elements

def generate_bom_data_pe (bom):
    '''read bom file and make data for pe3 doc'''
    m = WindowsPath(__file__)
    t = m.with_name("name.json")
    with open(t, "r", encoding="utf-8") as fh:
        name_json = load(fh) # download gbnk from file
    data_dev = list(map(change_quantity, union_same_element(read_bom_csv(bom)[0])))
    new_data_dev = list_new_pe(data_dev, name_json)
    finell_data_dev = add_empty_row(new_data_dev)
    return (finell_data_dev, read_bom_csv(bom)[1])


if __name__ == "__main__":
    #p = WindowsPath(__file__)
    #t = p.with_name('list.csv')
    #print('Path: ', p)
    #print('Path: ', t)
    #t = "list_biab100li.json"
    #with open(t, "r", encoding="utf-8") as fh:
    #    dev_json = load(fh) # download gbnk from file
        #print(dev_json)
    #key = "ГБНК.442259.025"
    #key = "ГБНК.566111.011"
    #if key in dev_json:
    #    print (dev_json[key])
    #t = "name.json"
    #with open(t, "r", encoding="utf-8") as fh:
    #    name_json = load(fh) # download gbnk from file
    #data_dev = listing_pe(key, 'VP100LI.csv')
    #new_data_dev = list_new_pe(data_dev, name_json)
    #finell_data_dev = add_empty_row(new_data_dev)
    #for i, d in enumerate(finell_data_dev):
    #    print(i, d)
        #print(name_json)
    #print(namedev_list(t, 'ГБНК.566111.011'))
    #print(namedev_list(t, 'ГБНК.674865.019'))
    #print(namedev_list(t))
    pass
