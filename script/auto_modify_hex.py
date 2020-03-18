# -*- coding: utf-8 -*-
# @Time    : 2020/3/18 11:56 上午
# @Author  : xiayuguo 
# @File    : auto_modify_hex.py
# @Software: PyCharm
import binascii
import openpyxl
import re


def read_config():
    configs = []
    wb = openpyxl.load_workbook('armor.xlsx')
    sheet = wb.active
    for row in list(sheet.rows)[1:]:
        config = {}
        if not row[0].value or not row[1].value:
            continue
        config['target'] = row[0].value.lower()
        config['value'] = row[1].value.lower()
        config['comment'] = row[2].value
        configs.append(config)
    return configs


def main_process():
    conf = read_config()
    print('load %s armor data to change' % len(conf))
    with open('../file/armor.am_dat', 'rb') as f:
        datas = f.read()
        a = binascii.b2a_hex(datas)
        a_str = a.decode()
        for armor in conf:
            # print(armor)
            if len(armor['target']) != len(armor['value']):
                print("[warning] value %s length not match %s!" % (armor['value'], armor['target']))
                continue
            if a_str.find(armor['target']) != -1:
                a_str = re.sub(armor['target'], armor['value'], a_str)
                print("%s has been successfully changed." % armor['comment'])
            else:
                print("cannot find armor %s, please check if the hex value is correct" % armor['comment'])
        a_byte = a_str.encode()
        new_a = binascii.a2b_hex(a_byte)

        with open('../output/armor.am_dat', 'wb') as f2:
            f2.write(new_a)
        print('success')

if __name__ == '__main__':
    main_process()
