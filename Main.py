import pandas as pd
import pyautogui
import time
from win10toast import ToastNotifier

# 初始化通知工具
toast = ToastNotifier()

deny = 0.2

# 读取Excel文件（请根据实际文件路径修改）
df = pd.read_excel('志愿信息.xlsx', dtype={
    '院校代码': str,
    '专业组': str,
    '代码1': str,
    '代码2': str,
    '代码3': str,
    '代码4': str,
    '代码5': str,
    '代码6': str,
    '是否服从调剂': str
})

def format_code(value):
    """处理专业代码格式，确保两位数"""
    try:
        return f"{int(value):02d}"
    except:
        return str(value).zfill(2)

# 显示初始通知
toast.show_toast(
    "志愿自动填报程序",
    "请在10秒内将光标定位到网页的第一个输入框（院校代码）...",
    duration=10,
    threaded=True
)

time.sleep(10)  # 等待用户定位光标

for index, row in df.iterrows():
    # 输入院校代码
    pyautogui.typewrite(row['院校代码'])
    time.sleep(deny)
    pyautogui.press('tab')
    time.sleep(deny)
    # 输入专业组
    pyautogui.typewrite(row['专业组'])
    time.sleep(deny)
    pyautogui.press('tab')
    time.sleep(deny)

    # 输入专业1-6
    for col in ['代码1', '代码2', '代码3', '代码4', '代码5', '代码6']:
        value = str(row[col]).strip()
        if pd.isna(row[col]) or value == '':
            pyautogui.press('tab')
            time.sleep(deny)
        else:
            pyautogui.typewrite(format_code(value))
            time.sleep(deny)
            pyautogui.press('tab')
            time.sleep(deny)
    time.sleep(deny)
    # 处理服从调剂
    adjust = str(row['是否服从调剂']).strip()
    if adjust == '是':
        pyautogui.press('space')
        time.sleep(deny)
        pyautogui.press('tab')
        time.sleep(deny)
        pyautogui.press('tab')
    elif adjust == '否':
        pyautogui.press('tab')
        time.sleep(deny)
        pyautogui.press('space')
        time.sleep(deny)
        pyautogui.press('tab')
    time.sleep(deny)

    # 发送进度通知
    toast.show_toast(
        "填报进度更新",
        f"志愿 {index+1}/{len(df)} 填写完成",
        duration=3,
        threaded=True
    )

time.sleep(2)
# 最终完成通知
toast.show_toast(
    "志愿自动填报程序",
    "所有志愿填写已完成！",
    duration=10,
    threaded=True
)
