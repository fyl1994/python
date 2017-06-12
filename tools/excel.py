#!/usr/bin/env python
# coding=utf-8

import sys
sys.path.append('./sony_excel_env/lib/python2.7/site-packages')
import openpyxl
from openpyxl.styles import Alignment
import datetime

def excel_xlsx(work_book):
    operator       = input("请输入运营商:")
    nv             = input("请输入NV:")
    operator_col    = -1
    modify_data     = -1
    sheet_id        = 0
    change_record   = "modify below value for"+operator+':\n'+nv+'='
    modify_data     = ''
    dms_id          = ''
    #遍历所有的sheet
    for sheet_name in (work_book.get_sheet_names()):
        print('\n','-----------------------\n','\n',sheet_name)
        #得到sheet
        sheet = work_book.get_sheet_by_name(sheet_name)
        #如果是历史页,且修改了数据,那么需要修改history
        if modify_data != '' and sheet_id == len(work_book.get_sheet_names())-1:
            history_row = sheet.max_row+1
            dms_id      = input("请输入DMS号:")
            record_alignment=Alignment(horizontal='left',vertical='bottom',text_rotation=0,wrap_text=True,shrink_to_fit=False,indent=0)
            sheet[history_row][4].alignment=record_alignment
            sheet[history_row][4].value=change_record
            if dms_id[0]!='D' or dms_id[0]!='d':
                dms_id="DMS"+dms_id
            if dms_id != '':
                sheet[history_row][3].value=dms_id
            date_now = datetime.datetime.now()
            year = date_now.strftime('%Y')
            month = date_now.strftime('%m')
            if month[0] == '0':
                month=month[1]
            day = date_now.strftime('%d')
            if day[0] == '0':
                day=day[1]            
            date_alignment=Alignment(horizontal='right',text_rotation=0,wrap_text=True,shrink_to_fit=False,indent=0)
            sheet[history_row][2].alignment=date_alignment
            sheet[history_row][2].value=year+'/'+month+'/'+day
            sheet.cell(row=history_row,column=3).number_format='yy/mm/dd@'
            work_book.guess_type=True
            return
        #遍历第一行,查找运营商 列
        for operator_col in range(sheet.max_column):
            if operator.lower() in (str(sheet.cell(row=1, column=operator_col+1).value).lower()):
                print("\n>>>>>>找到的运营商:\n",str(sheet.cell(row=1,column=operator_col+1).value))
                #遍历左8列,查找nv 行
                for temp_col in range(8):
                    for nv_row in range(sheet.max_row):
                        if nv.lower() in (str(sheet.cell(row=nv_row+1, column=temp_col+1).value).lower()) and nv_row > 0:
                            print("\n\n===找到NV:\n",str(sheet.cell(row=nv_row+1,column=temp_col+1).value))
                            #打印查找到的位置和数值
                            print("\nNV值为:\n",str(sheet.cell(row=nv_row+1,column=operator_col+1).value))
                            need_modify = input("\n是否需要修改[y/n]?")
                            if 'y' == need_modify:
                                modify_data = input("请输入要修改的值:\n")
                                change_record += str(modify_data)
                                sheet.cell(row = nv_row+1,column = operator_col+1).set_explicit_value(value = modify_data,data_type='n')
        sheet_id += 1;

def excel_save(workbook,file_path):
    date_temp=datetime.datetime.now().strftime("%Y%m%d")
    new_file_path=file_path[:-16]+date_temp[2:8]+file_path[-10:]
    workbook.save(filename=new_file_path)

if __name__ == '__main__':
    excel_filename = input("请输入表格绝对路径:")
    #区分Excel是xls还是xlsx
    if 'x' == excel_filename[-1] or 'X' == excel_filename[-1]:
        work_book = openpyxl.load_workbook(excel_filename)
        excel_xlsx(work_book)
        excel_save(work_book,excel_filename)

    if 's'== excel_filename[-1] or 'S' == excel_filename[-1]:
        print("暂不支持xls格式!!!")
