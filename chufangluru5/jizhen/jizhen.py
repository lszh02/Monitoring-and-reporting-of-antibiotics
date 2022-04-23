import pyautogui
import time
import xlrd
import pyperclip
import random
import datetime
import win32api
import pickle
import os


def mouseClick(img, clickTimes=1, lOrR="left"):
    # 定义鼠标事件
    # pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159
    while True:
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            break
        print("未找到匹配图片,0.5秒后重试")
        time.sleep(0.5)


def read_excel(path):
    file = xlrd.open_workbook(path)
    sheet = file.sheet_by_name("Sheet3")
    return sheet


def create_gender_dict():
    row = 1
    while row < worksheet.nrows:
        if '子宫' in worksheet.cell(row, 4).value \
                or '卵巢' in worksheet.cell(row, 4).value \
                or '阴道' in worksheet.cell(row, 4).value \
                or '妊娠' in worksheet.cell(row, 4).value \
                or '孕' in worksheet.cell(row, 4).value \
                or '乳腺' in worksheet.cell(row, 4).value:
            gender = "woman.png"
        elif '包皮' in worksheet.cell(row, 4).value \
                or '龟头' in worksheet.cell(row, 4).value \
                or '睾丸' in worksheet.cell(row, 4).value \
                or '前列腺' in worksheet.cell(row, 4).value \
                or '阴茎' in worksheet.cell(row, 4).value:
            gender = "man.png"
        else:
            gender = random.choice(["man.png", "woman.png"])
            i = 1
            while row - i > 0:
                if worksheet.cell(row, 1).value == worksheet.cell(row - i, 1).value \
                        and worksheet.cell(row, 2).value == worksheet.cell(row - i, 2).value:
                    gender = gender_dict.get(row - i)
                    break
                i += 1
        gender_dict.update({row: gender})

        ii = 1
        while row + ii < worksheet.nrows:
            if worksheet.cell_type(row + ii, 0) != 0:
                break
            ii += 1
        row += ii


def input_age():
    mouseClick("age.png")
    print("点击", "age.png")
    age = worksheet.cell(row_num, 2).value
    if '岁' in age:
        age = worksheet.cell(row_num, 2).value.split("岁")[0]
        pyperclip.copy(age)
        pyautogui.hotkey('ctrl', 'v')
        print(f"输入年龄:{age}岁")
    elif '月' in age:
        age = worksheet.cell(row_num, 2).value.split("月")[0]
        pyperclip.copy(age)
        pyautogui.hotkey('ctrl', 'v')
        mouseClick("age_year.png")
        mouseClick("age_month.png")
        print(f"输入年龄:{age}月")
    elif '天' in age:
        age = worksheet.cell(row_num, 2).value.split("天")[0]
        pyperclip.copy(age)
        pyautogui.hotkey('ctrl', 'v')
        mouseClick("age_year.png")
        mouseClick("age_day.png")
        print(f"输入年龄:{age}天")


def input_gender():
    gender = gender_dict.get(row_num)
    mouseClick(gender)
    print(f"选择性别：{gender.split('.')[0]}")


def input_moneyAnd_drugs_count():
    mouseClick("money.png")
    print("点击", "money.png")
    money = worksheet.cell(row_num, 3).value
    global row_x
    row_x = 1
    drugs_count = 1
    while row_num + row_x < worksheet.nrows:
        if worksheet.cell_type(row_num + row_x, 1) == 0:
            if worksheet.cell_type(row_num + row_x, 3):
                drugs_count += 1
                money += worksheet.cell(row_num + row_x, 3).value
            row_x += 1
        else:
            break
    pyperclip.copy(round(money, 2))
    pyautogui.hotkey('ctrl', 'v')
    print(f"输入处方金额:{money}")

    # 输入药品总数
    mouseClick("drugs_count.png")
    print("点击", "drugs_count.png")
    pyperclip.copy(drugs_count)
    pyautogui.hotkey('ctrl', 'v')
    print(f"输入药品数量:{drugs_count}")


def injection_or_not():
    inj_count = 0
    drug_name = worksheet.cell(row_num, 5).value
    if '注射' in drug_name or '狂犬病疫苗' in drug_name or '破伤风' in drug_name:
        inj_count += 1
    i = 1
    while row_num + i < worksheet.nrows:
        if worksheet.cell_type(row_num + i, 1) == 0:
            if worksheet.cell(row_num + i, 5).value != worksheet.cell(row_num + i - 1, 5).value:
                if '注射' in worksheet.cell(row_num + i, 5).value \
                        or '狂犬病疫苗' in worksheet.cell(row_num + i, 5).value \
                        or '破伤风' in worksheet.cell(row_num + i, 5).value:
                    inj_count += 1
            i += 1
        else:
            break
    if inj_count != 0:
        mouseClick("inj01.png")
        mouseClick("inj02.png")
        pyperclip.copy(inj_count)
        pyautogui.hotkey('ctrl', 'v')
        print("输入注射剂数量:", inj_count)


def input_diagnosis():
    mouseClick("diagnosis.png")
    print("点击", "diagnosis.png")
    diagnosis = worksheet.cell(row_num, 4).value
    # if worksheet.cell_type(row_num, 4) == 0:
    #     i = 1
    #     while worksheet.cell_type(row_num - i, 4) == 0:
    #         i += 1
    #     diagnosis = worksheet.cell(row_num - i, 4).value
    if '猫抓' in diagnosis or '猫咬' in diagnosis:
        diagnosis = '被其他哺乳动物咬伤或抓伤'
    if '癌' in diagnosis:
        diagnosis.replace('癌', '肿瘤')
    pyperclip.copy(f"{diagnosis}")
    pyautogui.hotkey('ctrl', 'v')
    mouseClick("search.png")
    print('请输入诊断！')
    while True:
        time.sleep(0.001)
        if win32api.GetKeyState(0x02) < 0:
            # up = 0 or 1, down = -127 or -128
            break


if __name__ == '__main__':
    # 打开文件，获取sheet3
    excel_path = r'D:\药事\抗菌药物监测\2022年\2022年3月'
    file_name = "2022年3月急诊处方点评.xls"
    worksheet = read_excel(rf"{excel_path}\{file_name}")
    print(rf"已打开工作表：{excel_path}\{file_name},获取sheet3")

    # 生成性别字典
    if not os.path.exists(rf'{excel_path}\急诊性别字典.txt'):
        with open(rf'{excel_path}\急诊性别字典.txt', "wb") as f:
            gender_dict = {}
            create_gender_dict()
            pickle.dump(gender_dict, f)
    else:
        print('\a')
        pyautogui.alert(text='性别字典已存在，新文件用旧字典会造成错误，请确认是断点续传！', title='请确认：', button='YES')
        # os.remove(rf'{excel_path}\急诊性别字典.txt')
        with open(rf'{excel_path}\急诊性别字典.txt', "rb") as f:
            gender_dict = pickle.load(f)
            print(gender_dict)

    pyautogui.alert(text='是否开始录入数据？', title='请确认：', button='YES')
    t1 = datetime.datetime.now()

    row_num = 43
    row_x = 1
    record = 0
    while row_num < worksheet.nrows:
        # 输入年龄
        input_age()

        # 选择性别
        input_gender()

        # 输入处方金额And药品总数
        input_moneyAnd_drugs_count()

        # 判断是否注射剂
        injection_or_not()

        # 输入诊断
        input_diagnosis()

        # 保存数据
        mouseClick("save.png")
        mouseClick("enter.png")
        # pyautogui.press('enter')
        print("点击", "保存数据")

        row_num += row_x
        record += 1
        print(f"————已遍历{row_num}行————已填报{record}条记录————")
    t2 = datetime.datetime.now()
    print(f"————共历时{t2 - t1}————")
