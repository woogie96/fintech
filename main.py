import os
import pandas as pd
import numpy as np

def FinTech_load():
    filenames = os.listdir(PATH)
    full_filenames = [os.path.join(PATH, filename)
                      for filename in filenames if ('FinTech' in filename and 'csv' in filename)]
    return full_filenames

def IT_select_load():
    filenames = os.listdir(PATH)
    full_filenames = [os.path.join(PATH, filename)
                      for filename in filenames if ('IT특강' in filename and 'csv' in filename)]
    return full_filenames

def FinProg_load():
    filenames = os.listdir(PATH)
    full_filenames = [os.path.join(PATH, filename)
                      for filename in filenames if ('금융IT' in filename and 'csv' in filename)]
    return full_filenames

def check_att(filename):
    name_min_dict = {}
    # csv 파일을 한줄씩 읽어옵니다.
    for each_line in open(filename, encoding='utf-16').readlines():
        # mins 가 포함된 라인만 확인합니다.
        if "mins" in each_line:
            # each_name에는 이름이나 학번이 들어가 있거나 둘다 포함되어있습니다.
            each_name = each_line.split("\t")[2]
            each_min = int(each_line.split("\t")[9].split()[0])
            # 30분 이상인 것만 출석으로 인정합니다.
            if each_min > 29:
                try:
                    name_min_dict[each_name] = each_min
                except:
                    # Dictionary의 중복을 방지합니다.
                    pass
    return name_min_dict

def make_att_excel(file_list,excel_name):
    excel = pd.read_excel(excel_name)
    for each_file in file_list:
        # 새 컬럼을 만들고 Nan을 채워넣습니다. 컬럼 이름은 각 날짜입니다.
        new_column_name = each_file[-8:-4]
        excel[new_column_name] = np.nan
        # 각 날짜마다 학생들의 출석이 기록된 Dictionary를 가져옵니다.
        checked_dict = check_att(each_file)
        for index in excel.index:
            # excel에 있는 학번이나 이름과 each_name이 일치하거나 포함될 경우 출석을 기록합니다.
            for each_name in checked_dict.keys():
                # target excel의 학번과 이름이 '학번', '이름'으로 표기된 경우입니다.
                try:
                    if str(excel.loc[index]['학번']) in each_name or excel.loc[index]['이름'] in each_name:
                        excel.loc[index, new_column_name] = 'O'
                except:
                # target excel의 학번과 이름이 'ID number', 'Fullname'으로 표기된 경우입니다.
                    if str(excel.loc[index]['ID number']) in each_name or excel.loc[index]['Fullname'] in each_name:
                        excel.loc[index, new_column_name] = 'O'
    return excel

if __name__ == '__main__':

    PATH = "C:/Users/fjdi7/Downloads/"
    os.chdir(PATH)

    FinTech_list = FinTech_load()
    FinTech_result = make_att_excel(FinTech_list,"FinTech_list.xlsx")
    FinTech_result.to_excel('FinTech_list_result.xlsx')

    FinProg_list = FinProg_load()
    FinProg_result = make_att_excel(FinProg_list,'FinProg_list.xlsx')
    FinProg_result.to_excel('FinFrog_list_result.xlsx')

    IT_select_list = IT_select_load()
    IT_select_result = make_att_excel(IT_select_list,'IT_SelectedTopics_list.xlsx')
    IT_select_result.to_excel('IT_select_result.xlsx')
