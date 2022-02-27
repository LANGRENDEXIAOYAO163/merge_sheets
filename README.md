# merge_sheets
合并多个Excel中的多个sheet到一个表格中

主要代码注释
###
# merged_file 目标文件名称
# tobe_merged 所有excel绝对路径LIST
def merge_excel_sheets(merged_file, tobe_merged):
    #using 1st excel format as standard
    std_fmt_file =  tobe_merged[0]
    #获取第一个文件的内容
    std_fmt_data = openpyxl.load_workbook(std_fmt_file)
    #获取所有Sheet名称LIST
    sheet_names =  std_fmt_data.sheetnames
   
    #initialize
    wb_merged_dict = dict()
    for name in sheet_names:
        wb_merged_dict[name] = [[]]
        #获取对应Sheet内容
        std_sheet = std_fmt_data[name]
        #first line (row) is table head
        #初始化各个sheet表头
        for head_col_idx in range(std_sheet.max_column):
            wb_merged_dict[name][0].append(std_sheet.cell(row=1,column=head_col_idx+1).value)
            logging.debug('wb_merged_dict[name][0]："%s"',wb_merged_dict[name][0])

    #merge all excel by sheet 
    #循环各个Excel
    for excel in tobe_merged:
        logging.debug('merging file:"%s" ...', excel)
        #获取对应表格
        ws_data=openpyxl.load_workbook(excel)
        #循环表格的各个Sheet
        for sheet_name in sheet_names:
            if sheet_name not in ws_data:
                logging.warn('WARNING: excel file: "%s" has no sheet named: "%s" will ignore it .', excel, sheet_name)
                continue
            #获取单个sheet
            sheet = ws_data[sheet_name]
            #初始化行
            sheet_row = 0
            #循环行
            for row_idx in range(sheet.max_row):
                if row_idx == 0:
                    continue
                sheet_row = sheet_row + 1
                #行数据
                row_data_list = []
                #循环每一行
                for col_idx in range(sheet.max_column):
                    #收集单行二维数组
                    row_data_list.append(sheet.cell(row=row_idx+1, column=col_idx+1).value)
                #将表格数据数据放到对应sheet中
                wb_merged_dict[sheet_name].append(row_data_list)
    #save data without format  (style)
    wb_merged = openpyxl.Workbook() #xlwt.Workbook(encoding = 'UTF-8')
    def_sheet = wb_merged.active
    wb_merged.remove(def_sheet)
    logging.debug("merged all file success , saving file:%s ...", merged_file)
    #循环表格LIST保存到目标表格中
    for sheet_name,sheet_data in wb_merged_dict.items():
        #创建对应Sheet
        sheet = wb_merged.create_sheet(title=sheet_name) # wb_merged.add_sheet(sheet_name)
        #循环取出对应的数据并且将数据放到对应位置
        for rowx in range(len(sheet_data)):
            for colx in range(len(sheet_data[rowx])):
                sheet.cell(row=rowx+1, column=colx+1, value=sheet_data[rowx][colx])
        logging.info('merged sheet:"%s" all files with data total row:%d .', sheet_name, len(sheet_data))
    wb_merged.save(merged_file)
def main():
    merged = "最终文件.xlsx"
    logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s', level=logging.DEBUG)
    dirname = "excel文件夹目录"
    tobe_merged = []
    for dn in os.listdir(dirname):
        if dn.endswith('.xlsx'):
            ##os.path.join 完全是拼接完整目录
            tobe_merged.append(os.path.join(dirname, dn))
        if len(tobe_merged) == 0:
            logging.error("not found xlsx file in dir:%s", dirname)
    merge_excel_sheets(merged, tobe_merged)
