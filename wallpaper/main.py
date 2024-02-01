# coding:utf-8
'''
本代码遵守GPLv3.0开源

特别鸣谢
代码修改自 Github @Asankilp (https://github.com/Asankilp/PyWallpaperEngine)
    (他的代码修改自 Bilibili @偶尔有点小迷糊 (https://b23.tv/BV1HZ4y1978a))
与
Github @BtbN (https://github.com/BtbN/FFmpeg-Builds)
'''

import json
import logging
import os
import time
import typing

import coloredlogs
import cv2
import win32con
import win32gui
import win32print

# 设置日志颜色
log_colors_config: dict[str, str] = {
    'DEBUG': 'white',
    'INFO': 'green',
    'WARNING': 'yellow',
    'ERROR': 'red',
    'CRITICAL': 'bold_red',
}

# 设置终端日志
coloredlogs.install(level='INFO', fmt='[%(levelname)s] [%(asctime)s]: %(message)s', colors=log_colors_config)
logging.info("日志设置成功, Wallpaper开始运行")

path: str = os.path.split(os.path.abspath(__file__))[0]
logging.debug(f"当前工作目录{path}")


# 获取真实的分辨率
def get_real_size() -> tuple[int, int]:
    # 获取真实的分辨率
    hDC: int = win32gui.GetDC(0)
    # 横向分辨率
    w: int = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)
    # 纵向分辨率
    h: int = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)
    return w, h


# 获取视频的分辨率
def get_video_size(path) -> tuple[float, float]:
    video: cv2.VideoCapture = cv2.VideoCapture(path)
    # 获取视频的宽度 (单位: 像素)
    w: float = video.get(cv2.CAP_PROP_FRAME_WIDTH)
    # 获取视频的高度 (单位: 像素)
    h: float = video.get(cv2.CAP_PROP_FRAME_HEIGHT)
    # 关闭视频文件
    video.release()
    return w, h


# 获取hide函数
def get_hide_func(process: os._wrap_close) -> typing.Callable[[int, None], None]:
    def hide(hwnd: int, hwnds: None) -> None:
        hdef: int = win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None)  # 枚举窗口寻找特定类
        if hdef != 0:
            workerw: int = win32gui.FindWindowEx(0, hwnd, "WorkerW", None)  # 找到hdef后寻找WorkerW
            win32gui.ShowWindow(workerw, win32con.SW_HIDE)  # 隐藏WorkerW
            logging.debug("已对窗口进行隐藏操作")
            logging.info("窗口设置完成")
            logging.info("动态壁纸已设定完成")
            try:
                process.close()  # 阻塞主程序, 直到ffplay异常退出
                logging.error("ffplay异常终止")
            except KeyboardInterrupt:
                logging.info("程序退出")
                exit()

    return hide


# 使用ffplay播放视频
def ffplay() -> os._wrap_close:
    try:
        with open(f"{path}\\config.json", "r", encoding="utf-8") as f:
            config: dict[str, str] = json.load(f)
    except IOError:
        logging.error("配置文件读取失败")
        with open(f"{path}\\config.json", "w", encoding="utf-8") as f:
            config = {"video": ""}
            # 保存
            json.dump(config, f, ensure_ascii=False, indent=4)

        logging.debug("配置文件创建成功")
        exit()
    except json.JSONDecodeError:
        logging.error("配置文件格式错误")
        exit()
    else:
        if (config["video"] == ""):
            logging.error("请先配置视频文件路径")
            exit()
        else:
            logging.info("配置文件读取成功")
    video = config["video"]

    # 自适应全屏, 防止黑边问题
    w, h = get_real_size()
    vw, vh = get_video_size(video)
    p = vw / vh
    if p <= w / h:
        dvh = h
        dvw = int(dvh * p)
        dx = int((w - dvw) / 2)
        dy = 0
    else:
        dvw = w
        dvh = int(dvw / p)
        dx = 0
        dy = int((h - dvh) / 2)

    # 无边框, 一直持续播放, 取消控制台输出
    return os.popen(f"{path}\\ffplay\\ffplay.exe {video} -noborder -left {dx} -top {dy} -x {dvw} -y {dvh} -loop 0 -loglevel quiet")


def display() -> None:
    logging.info("正在启动ffplay播放器播放视频...")
    process: os._wrap_close = ffplay()
    while not win32gui.IsWindowVisible(win32gui.FindWindow("SDL_app", None)):
        time.sleep(0.1)
    logging.info("ffplay播放器启动成功")

    progman: int = win32gui.FindWindow("Progman", "Program Manager")  # 寻找Progman
    logging.debug(f"已寻找到Progman窗口, 窗口句柄为{progman}")
    win32gui.SendMessageTimeout(progman, 0x52C, 0, 0, 0, 0)  # 发送0x52C消息
    logging.debug("已对Progman窗口发送0x52C消息")
    videowin: int = win32gui.FindWindow("SDL_app", None)  # 寻找ffplay播放窗口
    logging.debug(f"已寻找到ffplay播放器窗口, 窗口句柄为{videowin}")
    win32gui.SetParent(videowin, progman)  # 设置子窗口
    logging.debug("已将ffplay播放器窗口设置为Progman窗口的子窗口")
    win32gui.EnumWindows(get_hide_func(process), None)  # 枚举窗口, 回调hide函数


display()
