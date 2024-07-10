import math
import pandas as pd
import os
import os.path
import camelot
import tabula
import shutil

local_path = 'D:\\PDF转换'
result_path = 'D:\\PDF转换'
backup_path = 'D:\\PDF转换/备份'

is_debug = True

def convert_samsung(file):
    all_datas = []
    row_column = ['Code', ' Description', ' RD_Date', ' UoM', ' Unit_Price', ' Vendor_Code', ' Specification',
                  ' Receiver', ' Qty', ' Amount']

    tables = camelot.read_pdf(file, flavor='stream', pages='1-end')
    result_df = pd.DataFrame()
    for table in tables:
        df = table.df
        for index, row in df.iterrows():
            if row.loc[0] == 'Code':
                result_df = df.loc[index + 2:].reset_index(drop=True)
                break
        rows = result_df.shape[0]
        for index in range(math.floor(rows / 4)):
            Code = result_df.loc[index * 4][0]
            Description = result_df.loc[index * 4 + 1][1]
            RD_Date = result_df.loc[index * 4 + 1][2]
            UoM = result_df.loc[index * 4 + 1][3]
            Unit_Price = result_df.loc[index * 4 + 1][4]
            Vendor_Code = result_df.loc[index * 4 + 2][0] + result_df.loc[index * 4 + 3][0]
            Specification = result_df.loc[index * 4 +
                                          2][1] + result_df.loc[index * 4 + 3][1]
            Receiver = result_df.loc[index * 4 + 2][2]
            Qty = result_df.loc[index * 4 + 2][3]
            Amount = result_df.loc[index * 4 + 2][4]
            all_datas.append([Code, Description, RD_Date, UoM, Unit_Price,
                              Vendor_Code, Specification, Receiver, Qty, Amount])

    df = pd.DataFrame(all_datas, columns=row_column)
    result_file = result_path + '/' + os.path.basename(file)
    df.to_excel(result_file.replace('.pdf', '.xlsx'), index=False)


# convert_samsung(local_path + '/三星/' + 'PurchaseOrder_2112183519_L10HRV.pdf')
# exit(0)

def convert_guotai(file):
    row_column = ['Ref.Num', '物料编码/版本', '规格描述', '数量', '单位', '单价(未税)', '单价(含税)', '金额(含税)',
                  '交期', '备注']

    page_dfs = tabula.read_pdf(file, pages='all', lattice=True)
    result_df = None
    for page_df in page_dfs:
        if not isinstance(result_df, pd.DataFrame):
            result_df = page_df[page_df.iloc[:, 3] > 0]
        else:
            result_df = pd.concat([result_df, page_df[page_df.iloc[:, 3] > 0]])
    result_df.columns = row_column
    result_file = result_path + '/' + os.path.basename(file)
    result_df.to_excel(result_file.replace('.pdf', '.xlsx'), index=False)


def convert_lansi(file):
    page_dfs = tabula.read_pdf(file, pages='all', lattice=True)
    result_df = None
    for page_df in page_dfs:
        if not isinstance(result_df, pd.DataFrame):
            result_df = page_df[(page_df['序号'] != '本页小计') & (page_df['序号'] != '合计')]
        else:
            result_df = pd.concat([result_df, page_df[(page_df['序号'] != '本页小计') & (page_df['序号'] != '合计')]])
    result_file = os.path.join(result_path, os.path.basename(file))
    result_df.to_excel(result_file.replace('.pdf', '.xlsx'), index=False)


# convert_lansi('D:/PDF转换/蓝思/' + '蓝思科技-4500032964-20240624093638.pdf')
# exit(0)

def convert_tata(file):
    page_dfs = tabula.read_pdf(file, pages='all', lattice=True)
    result_df = None
    for page_df in page_dfs[1:]:
        if not isinstance(result_df, pd.DataFrame):
            result_df = page_df
        else:
            result_df = pd.concat([result_df, page_df])

    result_file = result_path + '/' + os.path.basename(file)
    result_df.to_excel(result_file.replace('.pdf', '.xlsx'), index=False)


companys = ['TATA', '蓝思', '三星', '国泰']
for company in companys:
    # 遍历下面所有pdf文件
    company_path = os.path.join(local_path, company)
    if os.path.exists(company_path):
        for file in os.listdir(company_path):
            fullpath_file = os.path.join(company_path, file)
            fullpath_backup = os.path.join(backup_path, company)
            if os.path.isfile(fullpath_file):
                filename, extension = os.path.splitext(file)
                if extension == '.pdf':
                    if company == '三星':
                        convert_samsung(fullpath_file)
                    elif company == '国泰':
                        convert_guotai(fullpath_file)
                    elif company == '蓝思':
                        convert_lansi(fullpath_file)
                    elif company == 'TATA':
                        convert_tata(fullpath_file)
            if not os.path.exists(fullpath_backup):
                os.makedirs(fullpath_backup)
            if os.path.isfile(os.path.join(fullpath_backup, file)):
                os.remove(os.path.join(fullpath_backup, file))
            if not is_debug:
                shutil.move(fullpath_file, fullpath_backup)
