import pyautogui
import time
import xlrd
import pyperclip
import random
import datetime



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


def start():
    mouseClick("st1.png", 2)
    for item in range(2, 10):
        mouseClick(f"st{item}.png")
        print(f"点击图片：st{item}.png")


def read_excel(excel_path):
    file = xlrd.open_workbook(excel_path)
    sheet = file.sheet_by_index(0)
    return sheet


if __name__ == '__main__':
    i = 1
    time.sleep(0.5)
    while True:
        print(i)
        if pyautogui.keyDown('shift'):
            print("双击", i)
        time.sleep(1)
        i += 1

