# coding:utf-8
'''
本代码遵守GPLv3.0开源

特别鸣谢
代码修改自 Github @Asankilp。(https://github.com/Asankilp/PyWallpaperEngine)
    (他的代码修改自 Bilibili @偶尔有点小迷糊。(https://b23.tv/BV1HZ4y1978a)(https://www.bilibili.com/video/BV1HZ4y1978a?share_source=copy_web&vd_source=6e8f2cf91fd19a35dcb657d1a79e2a83))
与
Bilibili @账号已注销。(https://www.bilibili.com/read/cv12718054)
    (他的代码修改自 Github @imniko。(https://github.com/imniko/SetDPI))
还有
Github @BtbN。(https://github.com/BtbN/FFmpeg-Builds)
'''

import json
import os
import threading
import time

import win32api
import win32con
import win32gui
import win32print


path = os.path.split(os.path.abspath(__file__))[0]


def get_real_resolution():
    # 获取真实的分辨率
    hDC = win32gui.GetDC(0)
    # 横向分辨率
    w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
    # 纵向分辨率
    h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
    return w, h


def get_screen_size():
    # 获取缩放后的分辨率
    w = win32api.GetSystemMetrics(0)
    h = win32api.GetSystemMetrics(1)
    return w, h


def getdpi():
    real_resolution = get_real_resolution()
    screen_size = get_screen_size()

    screen_scale_rate = round(real_resolution[0] / screen_size[0], 2)
    screen_scale_rate = screen_scale_rate * 100
    return screen_scale_rate


def hide(hwnd, hwnds):
    hdef = win32gui.FindWindowEx(
        hwnd, 0, "SHELLDLL_DefView", None)  # 枚举窗口寻找特定类
    if hdef != 0:
        workerw = win32gui.FindWindowEx(
            0, hwnd, "WorkerW", None)  # 找到hdef后寻找WorkerW
        win32gui.ShowWindow(workerw, win32con.SW_HIDE)  # 隐藏WorkerW
        while True:
            time.sleep(100)  # 进入循环防止壁纸退出


# 使用ffplay播放视频
def ffplay():
    with open(f"{path}\\config.json", encoding="utf-8") as config:
        config = json.load(config)
    video = config["video"]
    os.popen(
        f"{path}\\ffplay\\ffplay.exe {video} -noborder -fs -loop 0 -loglevel quiet")
    # 无边框、全屏、一直持续播放、取消控制台的输出


def display():
    ffplay()
    time.sleep(0.5)
    progman = win32gui.FindWindow("Progman", "Program Manager")  # 寻找Progman
    win32gui.SendMessageTimeout(progman, 0x52C, 0, 0, 0, 0)  # 发送0x52C消息
    videowin = win32gui.FindWindow("SDL_app", None)  # 寻找ffplay 播放窗口
    win32gui.SetParent(videowin, progman)  # 设置子窗口
    win32gui.EnumWindows(hide, None)  # 枚举窗口，回调hide函数


def back():
    print('缩放自动调回初始值')
    location = f"{path}\\SetDpi.exe"
    global userdpi
    userdpi = str(userdpi)
    win32api.ShellExecute(0, 'open', location, userdpi, '', 1)


time.sleep(2)
userdpi = getdpi()
print('当前系统缩放率为:', userdpi, '%', end='')
if userdpi == 100:
    print('缩放率无需调整')
    display()
    print("动态壁纸已设定完成")
else:
    print('正在调整为100%缩放率')
    location = f"{path}\\SetDpi.exe"
    win32api.ShellExecute(0, 'open', location, ' 100', '', 1)
    print('运行中 请等待')
    time.sleep(3)

    display_thread = threading.Thread(target=display)
    back_thread = threading.Thread(target=back)

    display_thread.start()
    time.sleep(1)
    back_thread.start()
