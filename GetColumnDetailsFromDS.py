############################################################
# Step 1)  Use Datasource object from the Document API
############################################################
import pandas as pd
import os
from tableaudocumentapi import Datasource

############################################################
# Step 2)  Open the .tds we want to inspect
############################################################
file = "abc.tds"
sourceTDS = Datasource.from_file(file)
############################################################
# Step 3) Generate the meaningful comments from messy data
############################################################

def splitter(x):
    """

    Parameters
    ----------
    x : TYPE
        Extract column Comment from an unformatted Field.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    try:
        y = x.split('<run>')[1].split('<')[0]
        return y
    except:
        return "NA"

############################################################
# Step 4)  Print out all of the fields and what type they are
############################################################
        
print('----------------------------------------------------------')
print('--- {} total fields in this datasource'.format(len(sourceTDS.fields)))
print('----------------------------------------------------------')
field_1=[]
datatype_1=[]
aggregation_1=[]
description_1=[]
for count, field in enumerate(sourceTDS.fields.values()):
    print('{:>4}: {} is a {}'.format(count+1, field.name, field.datatype))
    field_1.append(field.name)
    datatype_1.append(field.datatype)
    aggregation_1.append(field.default_aggregation)
    description_1.append(field.description)
    blank_line = False
    if field.calculation:
        print('      the formula is {}'.format(field.calculation))
        
        blank_line = True
    if field.default_aggregation:
        print('      the default aggregation is {}'.format(field.default_aggregation))
        blank_line = True
    if field.description:
        print('      the description is {}'.format(field.description))
     

    if blank_line:
        print('')
print('----------------------------------------------------------')
df = pd.DataFrame({'Column Names':field_1,'Data Type':datatype_1,'Aggregator':aggregation_1,'Description':description_1})
df['Description New'] = df['Description'].apply(splitter)
base = os.path.basename(file)
os.path.splitext(base)
filename = os.path.splitext(base)[0]
df.to_excel(filename + '.xlsx',index=False,encoding="utf-8")

