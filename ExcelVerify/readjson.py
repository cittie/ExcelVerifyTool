# -*- coding: utf-8 -*-

import json

def json_to_data(path):
    with open(path, 'r', encoding = 'utf-8') as data_file:
        data = json.load(data_file)
    
    return data