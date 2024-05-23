'''main generate PE, SP, TC'''
import argparse
from pathlib import WindowsPath
import frame
import readlist



if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('vp', help='enter full file name VP')
    parser.add_argument('--type', default = 'PE' , help ='choice type PE or SP or TC')
    parser.add_argument('--gbnk', default = None , help ='enter full number of gbnk')
    args = parser.parse_args()
    vp = WindowsPath(args.vp)
    gbnk = args.gbnk
    type_file = args.type
    name_and_family = ('', '', '', '', '', gbnk, '')

    finell_data_dev = readlist.generate_data_pe (vp, gbnk)

    for i, d in enumerate(finell_data_dev):
        print(i, d)

    COUNTLIST = 1 + (len(finell_data_dev) - 23)//25
    NAMETESTFILE = vp.with_name(name_and_family[5] + '_pe.xlsx')
    frame.create_workbook(name_and_family, type_file, COUNTLIST, NAMETESTFILE, finell_data_dev)
