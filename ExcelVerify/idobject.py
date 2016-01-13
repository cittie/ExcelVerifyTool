# -*- coding: utf-8 -*-
import os.path

class IDList:

    def __init__(self):
        self.content = []
        self.duplicate_list = []
        self.COLUMN_TITLE_LINE = 0
        self.COLUMN_CONTENT_LINE = 1
        self.is_duplicate = False
        self.pre_path = ""
        self.extension_name = ""
        self.logger = None
        
    def check_content(self, content):
        if content and content != "0" and content != "-1":
            if content not in self.content:
                self.content.append(content)
            else:
                self.is_duplicate = True
                self.duplicate_list.append(content)
                
    def import_from_current_column(self, sheet, column_index):
        for row_index in range(self.COLUMN_CONTENT_LINE, sheet.nrows):
            cell = sheet.cell(row_index, column_index)
            
            # Type 2 means numbers, type 3 means date, both are float.
            # Convert integer value in xlsx file to original str form.
            # Otherwise 1024 will be convert to "1024.0" as float.  
            cell_value = cell.value
            if cell.ctype in (2, 3) and int(cell_value) == cell_value:                
                cell_value = int(cell_value)
            content = str(cell_value)
            
            # Add each member in a group with separation '|'
            if '|' in content:
                content_splited = content.split('|')
                for content in content_splited:
                    self.check_content(content)
            else:
                self.check_content(content)

    def import_from_sheet(self, sheet, column_name_list, is_contained = False):
        for column_index in range(sheet.ncols):
            for column_name in column_name_list:
                if is_contained:    # Fuzzy matching.
                    if column_name in sheet.cell(self.COLUMN_TITLE_LINE, column_index).value:
                        self.import_from_current_column(sheet, column_index)
                else:               # Excatly matching.
                    if column_name == sheet.cell(self.COLUMN_TITLE_LINE, column_index).value:
                        self.import_from_current_column(sheet, column_index)

    def import_from_workbook(self, 
                             workbook, column_name_list,
                             sheet_name = None, 
                             check_duplicate = False, 
                             is_contained = False):

        if sheet_name:
            sheet = workbook.sheet_by_name(sheet_name)
            self.import_from_sheet(sheet, column_name_list, is_contained)
        else:
            for sheet_index in range(workbook.nsheets):
                sheet = workbook.sheet_by_index(sheet_index)
                self.import_from_sheet(sheet, column_name_list, is_contained)

        if check_duplicate:
            if self.is_duplicate:
                self.logger.log_status("target", "id", "duplicate")
                for content in self.duplicate_list:
                    self.logger.log_status("target", content, "is_duplicate")
            else:
                self.logger.log_status("target", "id", "no_duplicate")

        if self.content:
            self.logger.log_status("target", "id", "success")
        else:
            self.logger.log_status("target", "id", "not_exist")

    def import_from_workbook_with_string(self, 
                                         workbook, string,
                                         check_duplicate = False, 
                                         is_contained = False):

        # Import ALL cells which is BEGINNING with string
        # If is_contained is enable, import ALL cells CONTAINS the string
        
        for sheet_index in range(workbook.nsheets):
            sheet = workbook.sheet_by_index(sheet_index)
            for column_index in range(sheet.ncols):
                for row_index in range(self.COLUMN_CONTENT_LINE, sheet.nrows):
                    cell = sheet.cell(row_index, column_index)
                    if cell.ctype == 1:         # Type 1 means cell value is an Unicode string.
                        content = cell.value
                        if content not in self.content:
                            if len(content) >= len(string):
                                if is_contained:    # Fully matching
                                    if string in content:
                                        self.content.append(content)
                                else:               # Matching only the beginning string.
                                    if content[:len(string)] == string:
                                        self.content.append(content)
                        else:
                            self.is_duplicate = False
                            self.duplicate_list.append(content)

        if self.content:
            self.logger.log_status("target", "IDS", "success")
        else:
            self.logger.log_status("target", "IDS", "not_exist")

        if check_duplicate:
            if self.is_duplicate:
                self.logger.log_status("target", "IDS", "duplicate")
                for content in self.duplicate_list:
                    self.logger.log_status("target", content, "is_duplicate")
            else:
                self.logger.log_status("target", "IDS", "no_duplicate")

    def import_from_sheet_array_with_string(self, 
                                         workbook_array, string,
                                         check_duplicate = False, 
                                         is_contained = False):

        for sheet in workbook_array:
            for row in sheet:
                for cell in row:
                    content = cell
                    if string not in self.content:
                        if len(content) >= len(string):
                            if is_contained:    # Fully matching
                                if string in content:
                                    self.content.append(content)
                            else:               # Matching only the beginning string.
                                if content[:len(string)] == string:
                                    self.content.append(content)
                    else:
                        self.is_duplicate = False
                        self.duplicate_list.append(content)

        if self.content:
            self.logger.log_status("target", "IDS", "success")
        else:
            self.logger.log_status("target", "IDS", "not_exist")

        if check_duplicate:
            if self.is_duplicate:
                self.logger.log_status("target", "IDS", "duplicate")
                for content in self.duplicate_list:
                    self.logger.log_status("target", content, "is_duplicate")
            else:
                self.logger.log_status("target", "IDS", "no_duplicate")


    def compare_as_source(self, target_id_list_object):
        error_count = 0

        for item in target_id_list_object.content:
            if item and item not in self.content:
                error_count += 1
                self.logger.log_status("target", item, "not_found")

        if error_count == 0:
            self.logger.log_status("target", "ID", "matched")
        else:
            self.logger.log_status("other", error_count, "error_count")

    def check_files_exist(self):    # Check if related resource files exist.
        error_count = 0

        for filename in self.content:
            filename = self.pre_path + filename + self.extension_name
            if not os.path.isfile(filename):
                error_count += 1
                self.logger.log_status("file", filename, "not_exist")

        if error_count == 0:
            self.logger.log_status("file", "files", "all_exist")
        else:
            self.logger.log_status("other", error_count, "error_count")

        
#  Class define ends.