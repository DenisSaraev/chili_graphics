#!/usr/bin/python3.7
'''
This script convert sensor's data to graphic
'''

import os,glob
import pandas as pd
import xlsxwriter
import logging

#Logger configuration
logger=logging.getLogger()
logger.setLevel(logging.INFO)
formatter=logging.Formatter('%(asctime)s %(name)s %(levelname)s: %(message)s')
#Logger to file
fh=logging.FileHandler('/home/pi/projects/results/logs/Chili_graphics.log')
fh.setLevel(logging.INFO)
fh.setFormatter(formatter)
logger.addHandler(fh)
#Logger to console
ch=logging.StreamHandler()
ch.setLevel(logging.INFO)
ch.setFormatter(formatter)
logger.addHandler(ch)

logger.info('Script started')

#Plant info. We will use it for page's name
chili_dict={'SoilMoisture-1':'null',
    'SoilMoisture-2':'Jalapeno_2',
    'SoilMoisture-3':'Jalapeno_old_3',
    'SoilMoisture-4':'CarolinaReaper_4',
    'SoilMoisture-5':'Jalapeno_5',
    'SoilMoisture-6':'Tomato_6',}

#Excel file with data
path_to_xlsx='/home/pi/projects/results/chili_graphics.xlsx'
writer=pd.ExcelWriter(path_to_xlsx, engine='xlsxwriter')
workbook  = writer.book
#Sheet with graphics
graph_sheet = workbook.add_worksheet(name='GRAPH')
#Creating chart for soil moisture
chart_soil = workbook.add_chart({'type': 'line'})
#Settings for graphic soil moisture
chart_soil.set_title ({'name':'Dynamic of soil moisture'}) 
chart_soil.set_x_axis({'name':'DATE'}) 
chart_soil.set_y_axis({'name':'SOIL MOISTURE'})
chart_soil.set_size({'width': 1440, 'height': 640})
logger.info(f'Open xlsx file {path_to_xlsx}')

#Open files with sensor's data
for every_file in glob.glob(os.path.join('/home/pi/projects/results/','SoilMoisture-*')):
    logger.info(f'Opening csv file {every_file}')
    #Each plant will have own sheet in xlsx. This is name of sheets
    name=os.path.basename(every_file)
    name=os.path.splitext(name)[0]
    name=chili_dict[name]
    logger.info(f'File for plant: {name}')
    
    df=pd.read_csv(every_file,header=None,sep='|',names=['PLANT','DATE','SOIL MOISTURE'])
    logger.info(f'Converted to dataframe')
    df.to_excel(writer, sheet_name=name, index=False)
    worksheet=writer.sheets[name]
    
    #Format settings
    header_format=workbook.add_format({
        'bold':True,
        'text_wrap':False,
        'valign':'top',
        'fg_color':'#D7E4BC',
        'border':1})
    text_format=workbook.add_format({
        'bold':False,
        'text_wrap':False,
        'valign':'top',
        'border':1})
    date_format=workbook.add_format({
        'bold':False,
        'text_wrap':False,
        'valign':'top',
        'num_format':'dd/mm/yyyy hh:mm',
        'border':1})
    
    #Confirm settings to table
    worksheet.set_column('A:A', 8, text_format)
    worksheet.set_column('B:B', 19, date_format)
    worksheet.set_column('C:C', 18, text_format)
    #Confirm settings to headers    
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0,col_num,value,header_format)
    worksheet.autofilter('A1:C1')
    logger.info(f'Formating page completed')
    
    #We need to count measurements for catch full table
    num_rows=df.count()[0]
    logger.info(f'Raws in file {num_rows}')
    #Chart for this plant
    chart_soil.add_series({ 
        'name':f'{name}', #Chart name
        'categories':f'={name}!$B$2:$B${num_rows+1}', #X-axis with date. 'num_rows+1' because first raw is header
        'values':f'={name}!$C$2:$C${num_rows+1}', }) #Y-axis with soil moisture
    logger.info(f'Chart {name} added')
#Insert chart on first page and save document
graph_sheet.insert_chart('A1', chart_soil)
logger.info(f'Chart for soil moisture inserted on page')

#New chart with themperature
chart_themp = workbook.add_chart({'type': 'line'})
#Settings for this graphic
chart_themp.set_title ({'name':'Dynamic of themperature'}) 
chart_themp.set_x_axis({'name':'DATE'}) 
chart_themp.set_y_axis({'name':'THEMPERATURE,C'})
chart_themp.set_size({'width': 1440, 'height': 640})

path_themp='/home/pi/projects/results/Themperature.csv'
df=pd.read_csv(path_themp,header=None,sep='|',names=['SENSOR','DATE','THEMPERATURE'])
logger.info(f'Csv with themperature converted to dataframe')
df.to_excel(writer, sheet_name='Themperature', index=False)
worksheet=writer.sheets['Themperature']
    
#Confirm settings to table
worksheet.set_column('A:A', 8, text_format)
worksheet.set_column('B:B', 19, date_format)
worksheet.set_column('C:C', 18, text_format)
#Confirm settings to headers    
for col_num, value in enumerate(df.columns.values):
    worksheet.write(0,col_num,value,header_format)
worksheet.autofilter('A1:C1')
logger.info(f'Formating page with themperature completed')

#We need to count measurements for catch full table
num_rows=df.count()[0]
logger.info(f'Raws in file with themperature: {num_rows}')
#Chart for this page
chart_themp.add_series({ 
    'name':f'Themperature', #Chart name
    'categories':f'=Themperature!$B$2:$B${num_rows+1}', #X-axis with date. 'num_rows+1' because first raw is header
    'values':f'=Themperature!$C$2:$C${num_rows+1}', }) #Y-axis with themperature
logger.info(f'Chart with themperature added')
#Insert chart on first page and save document
graph_sheet.insert_chart('A34', chart_themp)
logger.info(f'Chart with themperature inserted on page')    
    
writer.save()
logger.info(f'File saved')
logger.info('Script finished')