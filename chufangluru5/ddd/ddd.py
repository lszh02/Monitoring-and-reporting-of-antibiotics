import pyautogui
import time
import xlrd
import pyperclip
import win32api


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
    sheet = file.sheet_by_index(1)
    return sheet


def input_drug_name():
    drug_name = worksheet.cell(row_num, 2).value

    pyperclip.copy(drug_name)
    # pyperclip.copy("头孢氨苄片")
    pyautogui.hotkey('ctrl', 'v')
    print("请输入第%d行药品名称" % (row_num + 1))
    while True:
        time.sleep(0.001)
        if win32api.GetKeyState(0x02) < 0:
            # up = 0 or 1, down = -127 or -128
            break
    mouseClick("search.png")
    print("点击", "search.png")
    mouseClick("find.png")
    print("点击", "find.png")


def input_drug_count():
    mouseClick("drugs_count.png")
    print("点击", "drugs_count.png")
    drugs_count = worksheet.cell(row_num, 4).value
    pyperclip.copy(drugs_count)
    pyautogui.hotkey('ctrl', 'v')
    print("输入药品总数:", drugs_count)


def input_drug_money():
    mouseClick("drugs_money.png")
    print("点击", "drugs_money.png")
    drugs_money = worksheet.cell(row_num, 5).value
    pyperclip.copy(round(drugs_money, 2))
    pyautogui.hotkey('ctrl', 'v')
    print("输入药品总数:", drugs_money)


def update_drug_dict():
    row = 1
    l1 = len(drug_dict)
    while row < worksheet.nrows:
        if worksheet.cell_type(row, 1) != 0:
            key = worksheet.cell(row, 1).value
            if key not in drug_dict:
                dep_name = input(f'{key}  未关联对应科室字典，请输入！')
                drug_dict[key] = dep_name  # 增加一条，更新字典
                print(f"{key}:{dep_name}")
        row += 1
    l2 = len(drug_dict)
    if l2 > l1:
        print('药品名称字典已更新%d条记录' % (l2-l1))


if __name__ == '__main__':
    # 打开文件，获取sheet页
    excel_path = r'D:\药事\抗菌药物监测\2022年'
    file_name = "住院抗菌药物使用情况查询（第四季度）2021.xls"
    worksheet = read_excel(rf"{excel_path}\{file_name}")
    print(rf"已打开工作表：{excel_path}\{file_name},获取sheet")

    # 定义、更新科室字典
    drug_dict = {'妇科门诊': 'dep_fuchan.png', '耳鼻喉科门诊': 'dep_erbihou.png'}
    update_drug_dict()

    row_num = 1
    while row_num < worksheet.nrows - 1:
        # 逐行读取数据
        mouseClick("drug_name.png")
        print("点击", "drug_name.png")

        # 输入药品名称
        input_drug_name()

        # 输入药品数量
        input_drug_count()

        # 输入药品金额
        input_drug_money()

        # 保存数据
        mouseClick("save.png")
        mouseClick("enter.png")
        print("点击", "保存数据")

        row_num += 1
        print("———————已填报%d条记录——————" % row_num)
