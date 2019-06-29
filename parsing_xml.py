# -*- coding: utf-8 -*-
"""
Created on Tue Jun 25 2019
@author: MinQiang
"""
import glob
import os
import datetime
import time
import configparser
import winreg
import sys

# import ElementTree to analysis XML
try:
    import xml.etree.cElementTree as ET
except ImportError:
    import xml.etree.ElementTree as ET


# get directory of desktop
def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, "Desktop")[0]
    # get configuration file

# read xml file
def traversal_dir_xml(path):
    config_name = 'conf.ini'
    if getattr(sys, 'frozen', False):
        #print("exe")
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        #print("script")
        application_path = os.path.dirname(__file__)
    config_path = os.path.join(application_path, config_name)
    # judge conf.ini is existence
    if os.path.exists(config_path):
        conf = configparser.ConfigParser()
        conf.read(config_path)
        secs = conf.sections()
        # get parameter in conf.ini
        dict1 = {}
        for sec in secs:
            dict1[conf.get(sec, 'version') + '_table'] = conf.get(sec, 'table')
            dict1[conf.get(sec, 'version') + '_para'] = conf.get(sec, 'para')
            # judge DV document is existence
        if os.path.exists(path):
            # get DV file
            f = glob.glob(path + '\\*' + '_FDDLTE.xml')
            # create a .txt file
            f_txt = open(
                get_desktop() + '\DV\DV_check_result_' + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + '.txt', 'a')
            for file_name in f:
                # open DV file xml
                tree = ET.ElementTree(file=file_name)
                root = tree.getroot()
                version = root.attrib['ver']
                table_name = 'N/A'
                dv_para = 'N/A'
                dv_value = 'N/A'
                if (version + '_table' in dict1.keys()) & (version + '_para' in dict1.keys()):
                    table_name = dict1[version + '_table']
                    para_name = dict1[version + '_para']
                    # query table_name
                    for child_of_root in root.iterfind('DBTABLE[@Table=' + '\"' + table_name + '\"' + ']'):
                        # print child_of_root.tag, child_of_root.attrib
                        for sub_child in child_of_root:
                            # print sub_child.tag, sub_child.attrib
                            # go through DV parameter
                            for sub_sub_child in sub_child.iterfind('FIELD[@name=' + '\"' + para_name + '\"' + ']'):
                                # print sub_sub_child.tag
                                dv_para = sub_sub_child.attrib['name']
                                dv_value = sub_sub_child.attrib['value']
                f_txt.write(file_name + ',' + version + ',' + table_name + ',' + dv_para + ',' + dv_value + '\n')
                # close .txt file
            f_txt.close()
        else:
            print('Please create a folder named DV on the desktop!')
            time.sleep(3)
    else:
        print('Missing configuration file named conf.ini!')
        time.sleep(3)

# Note that the path name mustn't contain Chinese
traversal_dir_xml(get_desktop() + '\DV')