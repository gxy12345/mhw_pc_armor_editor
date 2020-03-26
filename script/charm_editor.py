# -*- coding: utf-8 -*-
# @Time    : 2020/3/20 2:11 下午
# @Author  : xiayuguo 
# @File    : charm_editor.py
# @Software: PyCharm
import openpyxl
import binascii
import re


class Config(object):
    def __init__(self):
        self.config_name = ""
        self.origin_skill_id = "00"
        self.origin_skill_level = "00"
        self.target_skill_ids = ["00"] * 5
        self.target_skill_levels = ["00"] * 5

    @property
    def origin_hex_str(self):
        return "00"*15 + self.origin_skill_id + "00" + self.origin_skill_level + "00"*6

    @property
    def target_hex_str(self):
        hex_str = "00"*9
        # 如果没有指定技能3，则保持原有技能
        if self.target_skill_ids[2] == "00":
            self.target_skill_ids[2] = self.origin_skill_id
            self.target_skill_levels[2] = self.origin_skill_level
        for skill, level in zip(self.target_skill_ids, self.target_skill_levels):
            hex_str += skill + "00" + level
        return hex_str


# 护石自动修改
def read_skill():
    wb = openpyxl.load_workbook('charm.xlsx', data_only=True)
    skill_sheet = wb['skill']
    skill_data = {}
    for row in list(skill_sheet.rows)[1:]:
        skill_name = row[2].value
        skill_hex = row[4].value
        if len(skill_hex) == 1:
            skill_hex = '0' + skill_hex
        if skill_name == 'Unavailable':
            skill_hex = -1
        skill_data[skill_name] = skill_hex
    return skill_data


def read_config():
    skill_data = read_skill()
    wb = openpyxl.load_workbook('charm.xlsx', data_only=True)
    skill_sheet = wb['charm']
    charm_configs = []
    for row in list(skill_sheet.rows)[2:]:
        conf = Config()
        conf.config_name = row[0].value
        # 旧技能信息
        conf.origin_skill_id = skill_data[row[1].value]
        conf.origin_skill_level = '0' + str(row[2].value)
        # 写入新技能信息
        for i in range(5):
            if row[2*i+3].value:
                conf.target_skill_ids[i] = skill_data[row[2*i+3].value]
            if row[2*i+4].value:
                conf.target_skill_levels[i] = '0' + str(row[2*i+4].value)
        charm_configs.append(conf)
    return charm_configs


def main_process():
    conf = read_config()
    print('load %s armor data to change' % len(conf))
    with open('../file/armor.am_dat', 'rb') as f:
        byte_data = f.read()
        byte_s = binascii.b2a_hex(byte_data)
        hex_str = byte_s.decode()
        for armor in conf:
            print(armor.origin_hex_str)
            print(armor.target_hex_str)
            if re.search(armor.origin_hex_str, hex_str, flags= re.S | re.I):
                print(re.search(armor.origin_skill_id, hex_str, flags= re.S | re.I))
                hex_str = re.sub(armor.origin_hex_str, armor.target_hex_str, hex_str, flags= re.S | re.I)
                print("%s has been successfully changed." % armor.config_name)
            else:
                print("cannot find armor %s, please check if the hex value is correct" % armor.config_name)

            # if hex_str.find(armor.origin_hex_str) != -1:
            #     hex_str = re.sub(armor.origin_hex_str, armor.target_hex_str, hex_str)
            #     print("%s has been successfully changed." % armor.config_name)
            # else:
            #     print("cannot find armor %s, please check if the hex value is correct" % armor.config_name)
        hex_byte = hex_str.encode()
        new_byte_data = binascii.a2b_hex(hex_byte)

        with open('../output/armor.am_dat', 'wb') as f2:
            f2.write(new_byte_data)
        print('success')


if __name__ == '__main__':
    main_process()
