# -*- coding: utf-8 -*-

import xlrd
import configparser
import sys
import os.path
import customlog as clog

from idobject import IDList
from readxml import xml_to_sheet_array
from readjson import json_to_data

def create_object(COLUMN_TITLE_LINE, COLUMN_CONTENT_LINE, logger):
    obj = IDList()
    obj.COLUMN_TITLE_LINE = COLUMN_TITLE_LINE
    obj.COLUMN_CONTENT_LINE = COLUMN_CONTENT_LINE
    obj.logger = logger
    
    return obj

def get_config(config_file):
    config = configparser.ConfigParser()

    try:
        file = open(config_file, "r")
    except IOError:
        logs.log_status("file", config_file, "not_exist")
        return False
    else:
        file.close()
        config.read(config_file)
        return config

def import_as_workbook(filename):
    logs.log_status("file", filename, "import")

    try:
        workbook = xlrd.open_workbook(EXCEL_PATH + filename)
    except IOError:
        logs.log_status("file", filename, "not_exist")
        return False
    else:
        return workbook

def import_ids_source(filename):
    logs.log_status("file", filename, "import")

    try:
        workbook = xlrd.open_workbook(filename)
    except IOError:
        logs.log_status("file", filename, "not_exist")
        return False
    else:
        return workbook

def check_id(section):          # Check if target ids exist in source id group.
    
    cs = configs[section]
    
    source_filename = cs['source']
    check_duplicate = cs.getboolean('check_duplicate')     
    source_sheet = cs['source_sheet']
    source_title = cs['source_title'].split(' ')
    source_workbook = import_as_workbook(source_filename)

    source = create_object(COLUMN_TITLE_LINE, COLUMN_CONTENT_LINE, logs)
    source.import_from_workbook(source_workbook, source_title, source_sheet, check_duplicate)

    target_filenames = cs['target'].split(' ')
    target_fuzzy = cs.getboolean('target_fuzzy')
    target_title = cs['target_title'].split(' ')
    for target_filename in target_filenames:
        target_workbook = import_as_workbook(target_filename)
        target = create_object(COLUMN_TITLE_LINE, COLUMN_CONTENT_LINE, logs)
        target.import_from_workbook(target_workbook, target_title, is_contained = target_fuzzy)

        source.compare_as_source(target)

def check_ids(section):         # Check if IDS exists.
    
    cs = configs[section]
    
    ids_source = cs['source']
    eigen_string = cs['eigen_string']
    is_xml = cs.getboolean('is_xml')
    check_duplicate = cs.getboolean('check_duplicate')
    
    source = create_object(COLUMN_TITLE_LINE, COLUMN_CONTENT_LINE, logs)
    
    if is_xml:
        source_array = xml_to_sheet_array(IDS_PATH + ids_source)
        source.import_from_sheet_array_with_string(source_array, eigen_string, check_duplicate)
    else:
        source_workbook = import_ids_source(IDS_PATH + ids_source)
        source.import_from_workbook_with_string(source_workbook, eigen_string, check_duplicate)
    
    target_filenames = cs['target'].split(' ')
    for target_filename in target_filenames:
        target_workbook = import_as_workbook(target_filename)
        target = create_object(COLUMN_TITLE_LINE, COLUMN_CONTENT_LINE, logs)
        target.import_from_workbook_with_string(target_workbook, eigen_string)

        source.compare_as_source(target)    
    
def check_group(section):       # Check if group ids are exist in group sheet.
    
    cs = configs[section]
    
    source_filename = cs['source']
    check_duplicate = cs.getboolean('check_duplicate')
    group_sheet = cs['group_sheet']
    group_title = cs['group_title'].split(' ')
    target_title = cs['target_title'].split(' ')
    group_workbook = import_as_workbook(source_filename)

    group = create_object(COLUMN_TITLE_LINE, COLUMN_CONTENT_LINE, logs)
    group.import_from_workbook(group_workbook, group_title, group_sheet, check_duplicate)
    
    target = create_object(COLUMN_TITLE_LINE, COLUMN_CONTENT_LINE, logs)
    target.import_from_workbook(group_workbook, target_title)

    group.compare_as_source(target)

def check_file(section):        # Check if target file exists. 
                                # Will NOT check duplicate as file may be used for multi positions.
    cs = configs[section]

    source_filename = cs['source']
    source_sheet = cs['source_sheet']
    source_title = cs['source_title'].split(' ')
    source_workbook = import_as_workbook(source_filename)

    source = create_object(COLUMN_TITLE_LINE, COLUMN_CONTENT_LINE, logs)
    source.import_from_workbook(source_workbook, source_title, source_sheet)
    source.pre_path = PROJECT_PATH + cs['pre_path']
    source.extension_name = cs['extension_name']

    source.check_files_exist()

def get_corresponse_list(workbook, in_list, sheet_name, location_list):
    
    out_list = []    
    sheet = workbook.sheet_by_name(sheet_name)
    
    for in_object in in_list:
        for row_index in range(1, sheet.nrows):
            cell = sheet.cell(row_index, 0)
            
            if in_object == cell.value:
                for column_index in location_list:
                    out_object = sheet.cell(row_index, column_index)
                    if out_object.value and out_object not in out_list:
                        out_list.append(out_object.value)
     
    return out_list    

def get_mission_ids_require_bundle():
    
    mission_workbook = xlrd.open_workbook(EXCEL_PATH + "Mission.xlsx")
    mission_ids = create_object(COLUMN_TITLE_LINE, COLUMN_CONTENT_LINE, logs)
    mission_ids.import_from_workbook(mission_workbook, ["MissionID"], "Mission")
    
    return mission_ids

def check_autobundlelist(section):
    
    cs = configs[section]
    json_filename = cs['json_file']
    
    autodownloadlist_json_data = json_to_data(JSON_PATH + json_filename)
    mission_workbook = xlrd.open_workbook(EXCEL_PATH + "Mission.xlsx")
    monster_workbook = xlrd.open_workbook(EXCEL_PATH + "MonsterData.xlsx")
   
    used_mission_id_list = []
    # Python 3 only:
    #for package_name, package in autodownloadlist_json_data.items():
    # Python 2 only:
    for package_name, package in autodownloadlist_json_data.iteritems():
        
        if (package_name != "Default_Package" 
            and package_name != 'TT_PACKAGE'
            and package['filter_type'] == "missionid"):
            
            bundle_group_list = []
            mission_id_list = []
           
            for bundle_group in package['bundle_groups']:
                bundle_group_list.append(bundle_group)
                    
            for mission_id in package['filter']:
                mission_id_list.append(mission_id)
                # Store to ensure all extra missions are included.
                if mission_id not in used_mission_id_list:
                    used_mission_id_list.append(mission_id)
            
            quest_list = get_corresponse_list(
                mission_workbook, 
                mission_id_list, 
                "Mission",
                [6])
            
            scene_list = get_corresponse_list(
                mission_workbook, 
                quest_list, 
                "QuestConfig",
                range(1, 9, 2))
                
            scene_group_list = get_corresponse_list(
                monster_workbook,
                scene_list,
                "GroupInfo",
                [1])        
                
            tirgger_group_list = get_corresponse_list(
                mission_workbook, 
                quest_list, 
                "QuestConfig",
                range(2, 10, 2))
            
            trigger_id_list = get_corresponse_list(
                mission_workbook, 
                tirgger_group_list,
                "TriggerGroup",
                range(3, 6))
                
            enemy_id_list = get_corresponse_list(
                mission_workbook,                
                trigger_id_list,
                "EnemyTrigger",
                range(1, 27, 9))
                
            actor_id_list = get_corresponse_list(
                monster_workbook,
                enemy_id_list,
                "MonsterData",
                [1])
                
            monster_group_list = get_corresponse_list(
                monster_workbook,
                actor_id_list,
                "GroupInfo",
                [1])
            
            '''
            print(package_name)
            print(bundle_group_list)
            print(monster_group_list) 
            print(scene_group_list)
            '''
            used_bundle_list = scene_group_list + monster_group_list
            
            for bundle in used_bundle_list:
                if bundle not in bundle_group_list:
                    logs.log_status("file", package_name, "fail") 
                    logs.log_status("other", bundle, "bundle_miss")            
            
    # Get mission id and filter with key words.
    mission_ids_with_bundle = get_mission_ids_require_bundle()
    mission_ids_with_bundle.content = [
        mission_name for mission_name in mission_ids_with_bundle.content 
        if "Daily" in mission_name 
        or "Boss" in mission_name
        and "Test" not in mission_name]
    
    used_mission_ids = create_object(COLUMN_TITLE_LINE, COLUMN_CONTENT_LINE, logs)
    used_mission_ids.content = used_mission_id_list
    used_mission_ids.compare_as_source(mission_ids_with_bundle)  
                             
if __name__ == "__main__":
    logs = clog.CustomLog()
    config_file = "config.ini"      # Define configure file name.

    try:
        configs = get_config(config_file)
    except IOError:
        logs.log_status("file", config_file, "not_exist")
    else:
        logs.log_status("file", config_file, "import")

    if configs:
        #  Line number should be integers, add exception if needed.
        COLUMN_TITLE_LINE = int(configs['DEFAULT']['COLUMN_TITLE_LINE'])
        COLUMN_CONTENT_LINE = int(configs['DEFAULT']['COLUMN_CONTENT_LINE'])
        PROJECT_PATH = str(configs['DEFAULT']['PROJECT_PATH'])
        EXCEL_PATH = str(configs['DEFAULT']['EXCEL_PATH'])
        IDS_PATH = str(configs['DEFAULT']['IDS_PATH'])
        JSON_PATH = str(configs['DEFAULT']['JSON_PATH'])
        
        for section in configs.sections():
            logs.log_status("other", section, "section_start")
            
            check_type = configs[section]['check_type']

            if check_type == "id":
                check_id(section)
            elif check_type == "group":
               check_group(section)
            elif check_type == "file":
                check_file(section)
            elif check_type == "ids":
                check_ids(section)
            elif check_type == "json":
                check_autobundlelist(section)
            else:
                logs.log_status("other", section, "invalid_check_type")
                
            logs.log_status("other", section, "section_finished")

    logs.write_log("Report.txt")
