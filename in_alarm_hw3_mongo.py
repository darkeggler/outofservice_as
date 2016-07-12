# _*_ coding:utf-8 _*_
import json


json_file_name = 'readme.json'

with open(json_file_name, "r") as fp :
    d  = json.load(fp)