# -*- coding: utf-8 -*-

# This script will extract calculated fields and parameters
# from a Tableau workbook and output a CSV file with four columns:
# Calculation Name, Remote Name, Formula, and Comment
# The comment is from the calculated field when a // is used
# Comments should be placed before the calculation to be extracted properly


import xml.etree.ElementTree as ET
import os
import pandas as pd
from tableaudocumentapi import Workbook

# prompt user for twb file

file = "abc.twb"

# parse the twb file
tree = ET.parse(file)
root = tree.getroot()

worksheet_fields = {}
# list of calc's name, tableau generated name, and calculation/formula
fieldList = []
root = tree.getroot()

for worksheet in root.findall('.//worksheet'):
    if worksheet is None:
        continue
    else:
        worksheet_name = worksheet.attrib['name']
        for data in worksheet.findall('.//datasource-dependencies'):
            if data is None:
                continue
            else:
                for field in data.findall('.//column[@caption]'):
                    if field is None:
                        continue
                    else:
                        if field.find(".//calculation[@formula]") is None:
                            field_remote_name = field.attrib['name']
                            field_calculated_formula = ''
                            field_calculated_comment = ''
                            field_datatype = field.attrib['datatype']
                            field_custom_name = field.attrib['caption']
                            field_role = field.attrib['role']
                            field_type = field.attrib['type']
                            field_row = (
                                worksheet_name, field_remote_name, field_datatype, field_calculated_formula,
                                field_calculated_comment, field_custom_name, field_role, field_type)
                            fieldList.append(field_row)
                        else:
                            field_remote_name = field.attrib['name']
                            field_calculated_formula = field.find(".//calculation").attrib['formula']
                            calc_comment = ''
                            calc_formula = ''
                            for line in field_calculated_formula.split('\r\n'):
                                if line.startswith('//'):
                                    calc_comment = calc_comment + line + ' '
                                else:
                                    calc_formula = calc_formula + line + ' '
                            field_datatype = field.attrib['datatype']
                            field_custom_name = field.attrib['caption']
                            field_role = field.attrib['role']
                            field_type = field.attrib['type']
                            field_row = (
                                worksheet_name, field_remote_name, field_datatype, calc_formula, calc_comment,
                                field_custom_name, field_role, field_type)
                            fieldList.append(field_row)

# convert the list of calcs into a data frame
data = fieldList

data = pd.DataFrame(data, columns=['Worksheet Name', 'Remote Name', 'Data Type', 'Calculated Field Formula',
                                   'Calculated Field Comments', 'Custom Name', 'Role', 'Type'])

# remove duplicate rows from data frame
data = data.drop_duplicates(subset=None, keep='first', inplace=False)
print(data)
# export to csv
# get the name of the file

base = os.path.basename(file)
os.path.splitext(base)
filename = os.path.splitext(base)[0]
data.to_excel(filename + '.xlsx',index=False,encoding="utf-8")
