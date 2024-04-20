# coding:utf-8
"""
本代码遵守GPLv3.0开源

特别鸣谢
代码修改自 GitHub @Asankilp。(https://github.com/Asankilp/PyWallpaperEngine)
    (他的代码修改自 Bilibili @偶尔有点小迷糊。(https://b23.tv/BV1HZ4y1978a)(https://www.bilibili.com/video/BV1HZ4y1978a?share_source=copy_web&vd_source=6e8f2cf91fd19a35dcb657d1a79e2a83))
与
Bilibili @账号已注销。(https://www.bilibili.com/read/cv12718054)
    (他的代码修改自 GitHub @imniko。(https://github.com/imniko/SetDPI))
还有
GitHub @BtbN。(https://github.com/BtbN/FFmpeg-Builds)
"""

import atexit
import logging
import os
import sys
import time
import winreg

import coloredlogs
import cv2
import pystray
import win32con
import win32gui
import win32print
import yaml
from PIL import Image
from pystray import MenuItem, Menu
from ttkbootstrap.dialogs import Messagebox

exit_flag = False
# ffpaly_pid=0

# 设置日志颜色
log_colors_config = {
    'DEBUG': 'white',
    'INFO': 'green',
    'WARNING': 'yellow',
    'ERROR': 'red',
    'CRITICAL': 'bold_red',
}


# 退出检测
@atexit.register
def at_exit_fun():
    global exit_flag
    if not exit_flag:
        os.system('"taskkill /F /IM ffplay.exe"')
        # 获取当前解释器路径
        p = sys.executable
        # 启动新程序(解释器路径, 当前程序)
        os.execl(p, p, *sys.argv)
    else:
        return


# 退出
def stop_():
    global exit_flag
    exit_flag = True
    os.system('"taskkill /F /IM ffplay.exe"')
    os._exit(0)


# 设置终端日志
coloredlogs.install(level='INFO', fmt='[%(levelname)s] [%(asctime)s]: %(message)s', colors=log_colors_config)
logging.info("日志设置成功，Wallpaper开始运行")

path = os.path.dirname(os.path.realpath(sys.argv[0]))
logging.debug(f"当前工作目录{path}")


def get_real_resolution() -> tuple[int, int]:
    # 获取真实的分辨率
    h_dc: int = win32gui.GetDC(0)
    # 横向分辨率
    w: int = win32print.GetDeviceCaps(h_dc, win32con.DESKTOPHORZRES)
    # 纵向分辨率
    h: int = win32print.GetDeviceCaps(h_dc, win32con.DESKTOPVERTRES)
    return w, h


def hide(hwnd: int, _) -> None:
    hdef: int = win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None)  # 枚举窗口寻找特定类
    if hdef != 0:
        work_erw: int = win32gui.FindWindowEx(0, hwnd, "WorkerW", None)  # 找到hdef后寻找WorkerW
        win32gui.ShowWindow(work_erw, win32con.SW_HIDE)  # 隐藏WorkerW


def get_video_size(path):
    video = cv2.VideoCapture(path)
    # 获取视频的宽度（单位：像素）
    w = video.get(cv2.CAP_PROP_FRAME_WIDTH)
    # 获取视频的高度（单位：像素）
    h = video.get(cv2.CAP_PROP_FRAME_HEIGHT)
    # 关闭视频文件
    video.release()
    return w, h


# 使用ffplay播放视频
def ffplay() -> None:
    config_path = os.path.join(path, "config.yml")
    try:
        with open(config_path, encoding="utf-8") as f:
            config = yaml.load(f.read(), Loader=yaml.FullLoader)
    except IOError:
        logging.error("配置文件读取失败")
        with open(config_path, "w", encoding="utf-8") as f:
            config = {"video": "", "adaptive": True}
            # 保存
            yaml.dump(data=config, stream=f, allow_unicode=True)

        logging.debug("配置文件创建成功")
        Messagebox.show_info("配置文件已生成在%s路径下！" % config_path, title="WallPaper")
        stop_()
    else:
        if "video" in config and config["video"] == "":
            logging.error("请先配置视频文件路径")
            Messagebox.show_error("请先配置视频文件路径", title="WallPaper")
            stop_()
        else:
            logging.info("配置文件读取成功")
    video = config["video"]
    w, h = get_real_resolution()
    if "adaptive" in config:
        if config["adaptive"]:
            # 自适应全屏，防止黑边问题
            vw, vh = get_video_size(video)
            if vw == 0 or vh == 0:
                Messagebox.show_error("读取视频异常，请检查视频文件是否损坏，或是不支持的格式", title="WallPaper")
                stop_()
            p = vw / vh
            dvh = h
            dvw = int(dvh * p)
            dx = int((w - dvw) / 2)
            dy = 0
        else:
            dx = 0
            dy = 0
            dvw = w
            dvh = h
    else:
        Messagebox.show_error("配置文件中有项缺失，请尝试删除配置文件后重新运行", title="WallPaper")
        return

    # 无边框、一直持续播放、取消控制台的输出
    os.popen(
        f"{path}\\ffplay\\ffplay.exe -noborder -left {dx} -window_title \"WallPaper\""
        f" -top {dy} -x {dvw} -y {dvh} -loop 0 -i -loglevel quiet {video}"
    )


def display() -> None:
    logging.info("正在启动ffplay播放器播放视频...")
    ffplay()

    progman: int = win32gui.FindWindow("Progman", "Program Manager")  # 寻找Progman
    logging.debug(f"已寻找到Progman窗口，窗口句柄为{progman}")
    win32gui.SendMessageTimeout(progman, 0x52C, 0, 0, 0, 0)  # 发送0x52C消息
    logging.debug("已对Progman窗口发送0x52C消息")

    # 寻找ffplay 播放窗口
    flag = 0
    while not win32gui.IsWindowVisible(win32gui.FindWindow("SDL_app", None)):
        time.sleep(0.1)
        flag += 1
        if flag >= 50:
            Messagebox.show_error("ffpaly疑似启动失败，请重启测试，若依旧失败请修改配置文件测试！", title="WallPaper")
            stop_()
    logging.info("ffplay播放器启动成功！")
    videowin: int = win32gui.FindWindow("SDL_app", None)
    logging.debug(f"已寻找到ffplay播放器窗口，窗口句柄为{videowin}")
    win32gui.SetParent(videowin, progman)  # 设置子窗口
    logging.debug("已将ffplay播放器窗口设置为Progman窗口的子窗口")

    # 隐藏窗口
    win32gui.EnumWindows(hide, None)  # 枚举窗口，回调hide函数
    logging.debug("已对窗口进行隐藏操作")
    logging.info("窗口设置完成！")


os.system('"taskkill /F /IM ffplay.exe"')
display()
logging.info("动态壁纸已设定完成")

image = Image.open(os.path.join(path, "wallpaper.ico"))


def is_startup(name="WallPaper"):
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run",
                             winreg.KEY_SET_VALUE, winreg.KEY_ALL_ACCESS | winreg.KEY_WRITE | winreg.KEY_CREATE_SUB_KEY)
        value, _ = winreg.QueryValueEx(key, name)
        return True
    except:
        return False


def add_to_startup(name="WallPaper", file_path=""):
    if file_path == "":
        file_path = os.path.realpath(sys.argv[0])
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run",
                         winreg.KEY_SET_VALUE,
                         winreg.KEY_ALL_ACCESS | winreg.KEY_WRITE | winreg.KEY_CREATE_SUB_KEY)  # By IvanHanloth
    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, "\"" + file_path + "\"")
    winreg.CloseKey(key)
    Messagebox.show_info("已成功添加开机自启", title="WallPaper")


def reboot_ffplay():
    os.system('"taskkill /F /IM ffplay.exe"')
    display()


# 托盘菜单
menu = (MenuItem('添加自启动', lambda: add_to_startup()), MenuItem('重启ffplay', lambda: reboot_ffplay()),
        MenuItem('退出', lambda: stop_()), Menu.SEPARATOR, MenuItem("By:Xiaosu", None))
icon = pystray.Icon("name", title="WallPaper", icon=image, menu=menu)
icon.run()
