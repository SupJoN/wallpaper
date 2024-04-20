# coding:utf-8
"""
本代码遵守GPLv3.0开源

特别鸣谢
代码修改自 Github @Asankilp (https://github.com/Asankilp/PyWallpaperEngine)
    (他的代码修改自 Bilibili @偶尔有点小迷糊 (https://b23.tv/BV1HZ4y1978a))
与
Github @BtbN (https://github.com/BtbN/FFmpeg-Builds)
"""

import atexit
import json
import logging
import os
import subprocess
import sys
import time
import winreg

import coloredlogs
import pystray
import win32con
import win32gui
import win32print
import yaml
from PIL import Image
from pystray import MenuItem, Menu
from ttkbootstrap.dialogs import Messagebox

exit_flag: bool = False


# 退出检测
@atexit.register
def at_exit_fun():
    global exit_flag
    if not exit_flag:
        os.system('"taskkill /F /IM ffplay.exe"')
        # 获取当前解释器路径
        p: str = sys.executable
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
coloredlogs.install(level='INFO', fmt='[%(levelname)s] [%(asctime)s]: %(message)s')
logging.info("日志设置成功, Wallpaper开始运行")

path: str = os.path.dirname(os.path.realpath(sys.argv[0]))
config_path: str = os.path.join(path, "config.yml")

logging.debug(f"当前工作目录{path}")


# 读取配置文件
try:
    with open(config_path, encoding="utf-8") as f:
        config: dict = yaml.load(f.read(), Loader=yaml.FullLoader)
except FileNotFoundError:
    logging.warning("配置文件读取失败")
    with open(config_path, "w", encoding="utf-8") as f:
        config: dict = {
            "video": os.path.join(path, "TestWallpaperVideo", "nahida.webm"),
            "adaptive": True,
            "disable_audio": False,
            "pop_up_warnings": True
        }
        # 保存
        yaml.dump(data=config, stream=f, allow_unicode=True)

    logging.info("配置文件读取失败, \n新的配置文件已生成在%s路径下！")
    Messagebox.show_info("配置文件读取失败, \n新的配置文件已生成在%s路径下！" % config_path, title="WallPaper")
    stop_()

# 检查配置文件是否完整
if ("adaptive" in config and "disable_audio" in config and
        "video" in config and "pop_up_warnings" in config):
    
    if not isinstance(adaptive := config["pop_up_warnings"], bool):
        logging.error("警告弹窗的配置必须是布尔类型, 如果不知道如何修复, 请尝试删除配置文件")
        Messagebox.show_error("警告弹窗的配置必须是布尔类型, 如果不知道如何修复, 请尝试删除配置文件", title="WallPaper")
        stop_()
    
    if not isinstance(adaptive := config["adaptive"], bool):
        logging.error("自适应窗口的配置必须是布尔类型, 如果不知道如何修复, 请尝试删除配置文件")
        if config["pop_up_warnings"]:
            Messagebox.show_error("禁用音频配置必须是布尔类型, 如果不知道如何修复, 请尝试删除配置文件", title="WallPaper")
        stop_()
    if not isinstance(disable_audio := config["disable_audio"], bool):
        logging.error("禁用音频配置必须是布尔类型, 如果不知道如何修复, 请尝试删除配置文件")
        if config["pop_up_warnings"]:
            Messagebox.show_error("禁用音频配置必须是布尔类型, 如果不知道如何修复, 请尝试删除配置文件", title="WallPaper")
        stop_()
    if not os.path.exists(config["video"]) or not os.path.isfile(config["video"]):
        logging.error("视频文件路径不存在/不正确, 如果不知道如何修复, 请尝试删除配置文件")
        if config["pop_up_warnings"]:
            Messagebox.show_error("视频文件路径不存在/不正确, 如果不知道如何修复, 请尝试删除配置文件", title="WallPaper")
        
else:
    logging.error("配置文件中有项缺失, 请尝试删除配置文件后重新运行")
    if "pop_up_warnings" in config and config["pop_up_warnings"]:
        Messagebox.show_error("配置文件中有项缺失, 请尝试删除配置文件后重新运行", title="WallPaper")
    stop_()

logging.info("配置文件读取成功")


def get_real_size() -> tuple[int, int]:
    # 获取真实的分辨率
    h_dc: int = win32gui.GetDC(0)
    # 横向分辨率
    w: int = win32print.GetDeviceCaps(h_dc, win32con.DESKTOPHORZRES)
    # 纵向分辨率
    h: int = win32print.GetDeviceCaps(h_dc, win32con.DESKTOPVERTRES)
    return w, h


def hide(hwnd: int, _=None) -> None:
    hdef: int = win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None)  # 枚举窗口寻找特定类
    if hdef != 0:
        work_erw: int = win32gui.FindWindowEx(0, hwnd, "WorkerW", None)  # 找到hdef后寻找WorkerW
        win32gui.ShowWindow(work_erw, win32con.SW_HIDE)  # 隐藏WorkerW


# 获取视频的分辨率
def get_video_size(video: str) -> tuple[int, int]:
    # 使用ffprobe获取视频的分辨率
    process: subprocess.CompletedProcess = subprocess.run(f"\"{path}\\ffmpeg\\ffprobe.exe\" -v error "
                                                          f"-select_streams v:0 -show_entries stream=width,"
                                                          f"height -of json \"{video}\"", stdout=subprocess.PIPE,
                                                          stderr=subprocess.PIPE)
    # 错误处理
    if process.returncode != 0:
        logging.debug(f"ffprobe的终止代码为{process.returncode}")
        logging.error("请正确填写视频文件路径")
        if config["pop_up_warnings"]:
            Messagebox.show_error("请正确填写视频文件路径", title="WallPaper")
        stop_()
    data = json.loads(process.stdout)
    return data["streams"][0]["width"], data["streams"][0]["height"]


# 使用ffplay播放视频
def ffplay() -> subprocess.Popen:
    w, h = get_real_size()

    if config["adaptive"]:
        # 自适应全屏, 防止黑边问题
        vw, vh = get_video_size(config["video"])
        p: float = vw / vh
        if p <= w / h:
            dvh: int = h
            dvw: int = int(dvh * p)
            dx: int = int((w - dvw) / 2)
            dy: int = 0
        else:
            dvw: int = w
            dvh: int = int(dvw / p)
            dx: int = 0
            dy: int = int((h - dvh) / 2)
    else:
        # 固定全屏
        dvw: int = w
        dvh: int = h
        dx: int = 0
        dy: int = 0

    # 无边框, 一直持续播放, 取消控制台输出
    return subprocess.Popen(
        f'"{path}\\ffmpeg\\ffplay.exe" -noborder -left {dx} '
        f'-window_title "WallPaper" -top {dy} -x {dvw} -y {dvh} '
        f'-loop 0 -i -loglevel quiet {" -an" if config["disable_audio"] else ""} "{config["video"]}"',
        stdout=subprocess.PIPE, stderr=subprocess.PIPE
    )


def display() -> None:
    logging.info("正在启动ffplay播放器播放视频...")
    ffplay()

    progman: int = win32gui.FindWindow("Progman", "Program Manager")  # 寻找Progman
    logging.debug(f"已寻找到Progman窗口, 窗口句柄为0x{progman:08X}")
    win32gui.SendMessageTimeout(progman, 0x52C, 0, 0, 0, 0)  # 发送0x52C消息
    logging.debug("已对Progman窗口发送0x52C消息")

    # 寻找ffplay 播放窗口
    flag: int = 0
    while not win32gui.IsWindowVisible(win32gui.FindWindow("SDL_app", None)):
        time.sleep(0.1)
        flag += 1
        if flag >= 30:
            logging.error("ffpaly疑似启动失败, 请重启测试, 若依旧失败请修改配置文件测试！")
            if config["pop_up_warnings"]:
                Messagebox.show_error("ffpaly疑似启动失败, 请重启测试, 若依旧失败请修改配置文件测试！",
                                      title="WallPaper")
            stop_()
    logging.info("ffplay播放器启动成功！")
    video_win: int = win32gui.FindWindow("SDL_app", None)
    logging.debug(f"已寻找到ffplay播放器窗口, 窗口句柄为0x{video_win:08X}")
    win32gui.SetParent(video_win, progman)  # 设置子窗口
    logging.debug("已将ffplay播放器窗口设置为Progman窗口的子窗口")

    # 隐藏窗口
    win32gui.EnumWindows(hide, None)  # 枚举窗口, 回调hide函数
    logging.debug("已对窗口进行隐藏操作")
    logging.info("窗口设置完成！")


os.system('"taskkill /F /IM ffplay.exe"')
display()
logging.info("动态壁纸已设定完成")


def is_startup(name: str = "WallPaper") -> bool:
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run",
                             winreg.KEY_SET_VALUE, winreg.KEY_ALL_ACCESS | winreg.KEY_WRITE | winreg.KEY_CREATE_SUB_KEY)
        value, _ = winreg.QueryValueEx(key, name)
        return True
    except:
        return False


def add_to_startup(name: str = "WallPaper", file_path: str = "") -> None:
    if file_path == "":
        file_path = os.path.realpath(sys.argv[0])
    key: winreg.HKEYType = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run",
                                          winreg.KEY_SET_VALUE,
                                          winreg.KEY_ALL_ACCESS | winreg.KEY_WRITE | winreg.KEY_CREATE_SUB_KEY)  # By IvanHanloth
    winreg.SetValueEx(key, name, 0, winreg.REG_SZ, '"' + file_path + '"')
    winreg.CloseKey(key)
    Messagebox.show_info("已成功添加开机自启", title="WallPaper")


def reboot_ffplay() -> None:
    os.system('"taskkill /F /IM ffplay.exe"')
    display()


# 托盘菜单
menu: tuple = (MenuItem('添加自启动', lambda: add_to_startup()), MenuItem('重启ffplay', lambda: reboot_ffplay()),
               MenuItem('退出', lambda: stop_()), Menu.SEPARATOR, MenuItem("By:Xiaosu", None))

image: Image = Image.open(os.path.join(path, "wallpaper.ico"))

icon: pystray.Icon = pystray.Icon("name", title="WallPaper", icon=image, menu=menu)
icon.run()
