##########################################################################################################################################################################################
########################################     Pipeline Thermal Expansion & Displacements POSTPROCESSOR    ###############################################################################
########################################     Subject:    Abaqus FEA Postprocessing   ###############################################################################
########################################     Author :    Engr.Jesurobo Collins       #####################################################################################    #################################################################################################
########################################     Project:    Personal project            ##############################################################################################
########################################     Tools used: Python,Abaqus,xlsxwriter    ##############################################################################################
########################################     Email:      collins4engr@yahoo.com      ##############################################################################################
#########################################################################################################################################################################################

import sys,os
from abaqus import*
from abaqusConstants import*
from viewerModules import*
from math import*
import xlsxwriter
import glob
import numpy as np

# CHANGE TO CURRENT WORKING DIRECTORY
os.chdir('C:/temp/Pipeline Parametric studies/Expansion')

###CREATE EXCEL WORKBOOK, SHEETS AND ITS PROPERTIES####
execFile = 'Results.xlsx'
workbook = xlsxwriter.Workbook(execFile)
workbook.set_properties({
    'title':    'This is Abaqus postprocessing',
    'subject':  'Pipeline Thermal expansion analysis',   
    'author':   'Collins Jesurobo',
    'company':  'Personal Project',
    'comments': 'Created with Python and XlsxWriter'})

# Create a format to use in the merged range.
merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'yellow'})
SHEET1 = workbook.add_worksheet('Summary')
SHEET1.center_horizontally()
SHEET1.fit_to_pages(1, 1)
SHEET1.set_column(0,1,21)
SHEET1.set_column(2,3,19)
SHEET1.set_column(4,6,23)
SHEET1.merge_range('A1:G1', 'MAXIMUM AND MINIMUM DISPLACEMENT WITH CORRESPONDING WORST LOADCASE,WORST LOAD STEP AND NODE WHERE IT OCCURS',merge_format)
SHEET1.merge_range('A9:E9', 'WORST EXPANSION FOR EACH LOADCASE VERSUS WALL THICKNESS STUDIED',merge_format)

SHEET2 = workbook.add_worksheet('All_steps')
SHEET2.center_horizontally()
SHEET2.fit_to_pages(1, 1)
SHEET2.set_column(0,2,24)
SHEET2.merge_range('A1:F1', 'DISPLACEMENTS RESULTS FOR ALL PIPELINE NODES FOR EACH LOADCASE AND CORRESPONDING LOAD STEP',merge_format)


# defines the worksheet formatting (font name, size, cell colour etc.)
format_title = workbook.add_format()
format_title.set_bold('bold')
format_title.set_align('center')
format_title.set_align('vcenter')
format_title.set_bg_color('#F2F2F2')
format_title.set_font_size(10)
format_title.set_font_name('Arial')
format_table_headers = workbook.add_format()
format_table_headers.set_align('center')
format_table_headers.set_align('vcenter')
format_table_headers.set_text_wrap('text_wrap')
format_table_headers.set_bg_color('#F2F2F2')
format_table_headers.set_border()
format_table_headers.set_font_size(10)
format_table_headers.set_font_name('Arial')

###WRITING THE TITLES TO SHEET1,SHEET2###
SHEET1.write_row('B2',['U1 - Expansion(mm)','U2(mm)','U3(mm)','WorstLoadcase','LoadStep','Node'],format_title)
SHEET1.write('A3', 'Max value',format_title)
SHEET1.write('A4', 'Min value',format_title)
SHEET1.write('A5', 'Absolute Max value',format_title)
SHEET1.write_row('A10', ['Loadcase','Thickness(mm)','Max_Expansion(mm)','Min_Expansion(mm)','Abs_Expansion(mm)'],format_title)

SHEET2.write_row('A2',['Loadcase','LoadStep','Node','U1(mm)','U2(mm)','U3(mm)'],format_title)

###LOOP THROUGH THE ODBs, LOOP THROUGH EACH STEPS AND EXTRACT DISPLACEMENT RESULTS FOR ALL PIPELINE NODES###
def output1():
        row=1
        col=0
        for i in glob.glob('*.odb'):     # loop  to access all odbs in the folder
                odb = session.openOdb(i) # open each odb
                step = odb.steps.keys()  # probe the content of the steps object in odb, steps object is a dictionary, so extract the step names with keys()
                section = odb.rootAssembly.instances['PART-1-1'].nodeSets['PIPELINE'] # extract section for pipeline nodeset
                ###DEFINE RESULT OUTPUT####
                for k in range(len(step)):
                        U = odb.steps[step[k]].frames[-1].fieldOutputs['U'].getSubset(region=section).values   # results for all displacements U 
                        for disp in U:
                                U1 = disp.data[0]            # extract U1 (axial) displacements from all the odbs and loadcases
                                U2 = disp.data[1]            # extract U2 (lateral) displacements from all the odbs and loadcases
                                U3 = disp.data[2]            # extract U3 (vertical) displacements from all the odbs and loadcases
                                n1 = disp.nodeLabel          # extract node numbers 
                                ### WRITE OUT MAIN RESULT OUTPUT####
                                SHEET2.write(row+1,col,i.split('.')[0],format_table_headers)  # loadcases
                                SHEET2.write(row+1,col+1,step[k],format_table_headers)        # steps in odb
                                SHEET2.write(row+1,col+2,n1,format_table_headers)             # all nodes in the pipeline
                                SHEET2.write(row+1,col+3,U1*1000,format_table_headers)        # displacements in axial direction, in mm        
                                SHEET2.write(row+1,col+4,U2*1000,format_table_headers)        # displacements in lateral direction, in mm
                                SHEET2.write(row+1,col+5,U3*1000,format_table_headers)        # displacements in vertical direction,in mm
                                row+=1
output1()

### GET THE MAXIMUM AND MINIMUM, AND ABSOLUTE MAXIMUM VALUES AND WRITE THM INTO SUMMARY SHEET(SHEET1) 
def output2():
        SHEET1.write('B3', '=ROUND(max(All_steps!D3:D100000),1)',format_table_headers)   # maximum displacement in axial direction-max(U1)
        SHEET1.write('C3', '=ROUND(max(All_steps!E3:E100000),1)',format_table_headers)   # maximum displacement in lateral direction-max(U2)
        SHEET1.write('D3', '=ROUND(max(All_steps!F3:F100000),1)',format_table_headers)   # maximum displacement in vertical direction-max(U3)
        SHEET1.write('B4', '=ROUND(min(All_steps!D3:D100000),1)',format_table_headers)   # minimum displacement in axial direction-min(U1)
        SHEET1.write('C4', '=ROUND(min(All_steps!E3:E100000),1)',format_table_headers)   # minimum displacement in lateral direction-min(U2)
        SHEET1.write('D4', '=ROUND(min(All_steps!F3:F100000),1)',format_table_headers)   # minimum displacement in vertical direction-min(U3)
        SHEET1.write('B5','=IF(ABS(B3)>ABS(B4),ABS(B3),ABS(B4))',format_table_headers) # absolute maximum U1
        SHEET1.write('C5','=IF(ABS(C3)>ABS(C4),ABS(C3),ABS(C4))',format_table_headers) # absolute maximum U2
        SHEET1.write('D5','=IF(ABS(D3)>ABS(D4),ABS(D3),ABS(D4))',format_table_headers) # absolute maximum U3

        ### WORST LOADCASE AND LOADSTEP CORRESPONDING TO MAXIMUM AND MINIMUM EXPANSION VALUES
        SHEET1.write('E3','=INDEX(All_steps!A3:A100000,MATCH(MAX(All_steps!D3:D100000),All_steps!D3:D100000,0))',format_table_headers)
        SHEET1.write('F3','=INDEX(All_steps!B3:B100000,MATCH(MAX(All_steps!D3:D100000),All_steps!D3:D100000,0))',format_table_headers)
        SHEET1.write('G3','=INDEX(All_steps!C3:C100000,MATCH(MAX(All_steps!D3:D100000),All_steps!D3:D100000,0))',format_table_headers)
        SHEET1.write('E4','=INDEX(All_steps!A3:A100000,MATCH(MIN(All_steps!D3:D100000),All_steps!D3:D100000,0))',format_table_headers)
        SHEET1.write('F4','=INDEX(All_steps!B3:B100000,MATCH(MIN(All_steps!D3:D100000),All_steps!D3:D100000,0))',format_table_headers)
        SHEET1.write('G4','=INDEX(All_steps!C3:C100000,MATCH(MIN(All_steps!D3:D100000),All_steps!D3:D100000,0))',format_table_headers)
output2()

### LOADCASES
def output3():
        row=0
        col=0
        for LC in glob.glob('*.odb'):
                SHEET1.write(row+10,col,LC.split('.')[0],format_table_headers)
                row+=1

# WRITE THE COLUMN FOR WALL THICKNESSES THAT WAS USED IN THE PARAMETRIC STUDIES
Thick_data = [15.9,19.1,22.3,25.1,27.1,30.2]         # varied thickness in mm
SHEET1.write_column('B11',Thick_data,format_table_headers)

### WORST EXPANSION VALUES                

SHEET1.write('C11','{=MAX(IF(All_steps!A3:A100000=Summary!A11, All_steps!D3:D100000))}',format_table_headers)
SHEET1.write('C12','{=MAX(IF(All_steps!A3:A100000=Summary!A12, All_steps!D3:D100000))}',format_table_headers)
SHEET1.write('C13','{=MAX(IF(All_steps!A3:A100000=Summary!A13, All_steps!D3:D100000))}',format_table_headers)
SHEET1.write('C14','{=MAX(IF(All_steps!A3:A100000=Summary!A14, All_steps!D3:D100000))}',format_table_headers)
SHEET1.write('C15','{=MAX(IF(All_steps!A3:A100000=Summary!A15, All_steps!D3:D100000))}',format_table_headers)
SHEET1.write('C16','{=MAX(IF(All_steps!A3:A100000=Summary!A16, All_steps!D3:D100000))}',format_table_headers)

SHEET1.write('D11', '{=MIN(IF(All_steps!A3:A100000=Summary!A11,All_steps!D3:D100000))}',format_table_headers)
SHEET1.write('D12','{=MIN(IF(All_steps!A3:A100000=Summary!A12, All_steps!D3:D100000))}',format_table_headers)
SHEET1.write('D13','{=MIN(IF(All_steps!A3:A100000=Summary!A13, All_steps!D3:D100000))}',format_table_headers)
SHEET1.write('D14','{=MIN(IF(All_steps!A3:A100000=Summary!A14, All_steps!D3:D100000))}',format_table_headers)
SHEET1.write('D15','{=MIN(IF(All_steps!A3:A100000=Summary!A15, All_steps!D3:D100000))}',format_table_headers)
SHEET1.write('D16','{=MIN(IF(All_steps!A3:A100000=Summary!A16, All_steps!D3:D100000))}',format_table_headers)

SHEET1.write('E11','=IF(ABS(C11)>ABS(D11),ABS(C11),ABS(D11))',format_table_headers) 
SHEET1.write('E12','=IF(ABS(C12)>ABS(D12),ABS(C12),ABS(D12))',format_table_headers) 
SHEET1.write('E13','=IF(ABS(C13)>ABS(D13),ABS(C13),ABS(D13))',format_table_headers) 
SHEET1.write('E14','=IF(ABS(C14)>ABS(D14),ABS(C14),ABS(D14))',format_table_headers) 
SHEET1.write('E15','=IF(ABS(C15)>ABS(D15),ABS(C15),ABS(D15))',format_table_headers) 
SHEET1.write('E16','=IF(ABS(C16)>ABS(D16),ABS(C16),ABS(D16))',format_table_headers)


# CREATE A PLOT OF EXPANSION VERSUS PIPE WALL THICKNESS
chart = workbook.add_chart({'type': 'line'})

# Add a series to the chart.
chart.add_series({
        'name': 'Thermal Expansion',
        'categories':'=Summary!$B$11:$B$16',            #Thickness in x-axis
        'values': '=Summary!$E$11:$E$16',
        'line':{'color':'blue'}})                       #Expansion in y-axis
chart.set_title({'name': 'Thermal Expansion',})
chart.set_x_axis({'name': 'Pipeline Wall Thickness(mm)',})
chart.set_y_axis({'name': 'Thermal Expansion(mm)',})
chart.set_style(9)

# Insert the chart into the worksheet.
SHEET1.insert_chart('F8', chart)
output3()

# closes the workbook once all data is written
workbook.close()

# opens the resultant spreadsheet
os.startfile(execFile)

# parameteric study completed









