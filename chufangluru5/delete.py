import pyautogui
import time



def mouseClick(img, clickTimes=1, lOrR="left"):
    # 定义鼠标事件
    # pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159
    while True:
        location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
        if location is not None:
            pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            break
        print("未找到匹配图片,0.1秒后重试")
        time.sleep(0.1)


if __name__ == '__main__':
    while True:
        mouseClick('delete_enter.png')
        mouseClick('enter.png')
        mouseClick('enter.png')
