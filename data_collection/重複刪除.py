import os
import shutil

frist_path = 'C:\\Users\\GIGABYTE\\桌面\\VScode_Code\\天氣預測資料'
frist_list = os.listdir(frist_path)
for case_name in frist_list:
    if("py" in case_name):
        continue
    dir_path = os.path.join(frist_path, case_name)
    dir_list = os.listdir(dir_path)
    for dir_name in dir_list:
        if("py" in dir_name):
            continue
        file_path = os.path.join(dir_path, dir_name)
        file_list = os.listdir(file_path)
        for file_name in file_list:
            if("xlsx" in file_name):
                continue
            test = file_name.split('(')
            if(len(test) > 1):
                os.remove(os.path.join(file_path, file_name))
                print(file_name)