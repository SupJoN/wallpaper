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
import logging
import coloredlogs
import os
import time
import cv2
import win32con
import win32gui
import win32print

# 设置日志颜色
log_colors_config = {
    'DEBUG': 'white',
    'INFO': 'green',
    'WARNING': 'yellow',
    'ERROR': 'red',
    'CRITICAL': 'bold_red',
}

# 设置终端日志
coloredlogs.install(level='INFO', fmt='[%(levelname)s] [%(asctime)s]: %(message)s', colors=log_colors_config)
logging.info("日志设置成功，Wallpaper开始运行")

path = os.path.split(os.path.abspath(__file__))[0]
logging.debug(f"当前工作目录{path}")

def get_real_resolution() -> tuple[int, int]:
    # 获取真实的分辨率
    hDC: int = win32gui.GetDC(0)
    # 横向分辨率
    w: int = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
    # 纵向分辨率
    h: int = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
    return w, h

def hide(hwnd: int, hwnds: None) -> None:
    hdef: int = win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None)  # 枚举窗口寻找特定类
    if hdef != 0:
        workerw: int = win32gui.FindWindowEx(0, hwnd, "WorkerW", None)  # 找到hdef后寻找WorkerW
        win32gui.ShowWindow(workerw, win32con.SW_HIDE)  # 隐藏WorkerW
        while True:
            time.sleep(100)  # 进入循环防止壁纸退出

def get_video_size(path):
    video = cv2.VideoCapture(path)
    # 获取视频的宽度（单位：像素）
    w = video.get(cv2.CAP_PROP_FRAME_WIDTH)
    # 获取视频的高度（单位：像素）
    h = video.get(cv2.CAP_PROP_FRAME_HEIGHT)
    # 关闭视频文件
    video.release()
    return w,h

# 使用ffplay播放视频
def ffplay() -> None:
    with open(f"{path}\\config.json", encoding="utf-8") as config:
        config: dict[str, str] = json.load(config)
    video = config["video"]
    w,h=get_real_resolution()
    
    # 自适应全屏，防止黑边问题
    vw,vh=get_video_size(video)
    p=vw/vh
    dvh=h
    dvw=dvh*p
    dvh=int(dvh)
    dvw=int(dvw)

    dx=(w-dvw)/2
    dy=(h-dvh)/2
    dx=int(dx)
    dy=int(dy)

    os.popen(f"{path}\\ffplay\\ffplay.exe {video} -noborder -left {dx} -top {dy} -x {dvw} -y {dvh} -loop 0  -loglevel quiet")
    # 无边框、一直持续播放、取消控制台的输出


def display() -> None:
    logging.info("正在启动ffplay播放器播放视频...")
    ffplay()
    while True:
        if win32gui.IsWindowVisible(win32gui.FindWindow("SDL_app", None)):
            break
        time.sleep(0.1)
    logging.info("ffplay播放器启动成功！")
    
    progman: int = win32gui.FindWindow("Progman", "Program Manager")  # 寻找Progman
    logging.debug(f"已寻找到Progman窗口，窗口句柄为{progman}")
    win32gui.SendMessageTimeout(progman, 0x52C, 0, 0, 0, 0)  # 发送0x52C消息
    logging.debug("已对Progman窗口发送0x52C消息")
    videowin: int = win32gui.FindWindow("SDL_app", None)  # 寻找ffplay 播放窗口
    logging.debug(f"已寻找到ffplay播放器窗口，窗口句柄为{videowin}")
    win32gui.SetParent(videowin, progman)  # 设置子窗口
    logging.debug("已将ffplay播放器窗口设置为Progman窗口的子窗口")
    win32gui.EnumWindows(hide, None)  # 枚举窗口，回调hide函数
    logging.debug("已对窗口进行隐藏操作")
    logging.info("窗口设置完成！")

display()
logging.info("动态壁纸已设定完成")
