import threading
import time
import tkinter
import tkinter as tk
import traceback
from threading import Thread
from tkinter import *
from tkinter import ttk, colorchooser
from tkinter.filedialog import askopenfilename
from tkinter.scrolledtext import ScrolledText
from sqlitedict import SqliteDict
import keyboard as keyboard
import pythoncom
from PIL import Image, ImageTk
from markdown2 import Markdown
from TkHtmlView import TkHtmlView
import sys, os
from pydmdll import DM


class Main_window(object):
    """主窗体"""

    def __init__(self, title, icon, alpha, topmost, bg, width, height, width_adjust, higth_adjust):
        self.root = tk.Tk()
        self.title = title
        self.icon = icon
        self.root.title(title)
        self.root.wm_attributes('-alpha', alpha)
        self.root.wm_attributes('-topmost', topmost)
        self.root.configure(bg=bg)
        self.bg = bg
        self.root.iconbitmap(icon)
        self.width = width
        self.height = height
        self.width_adjust = width_adjust
        self.higth_adjust = higth_adjust

        # 文件执行目录
        self.exe_dir_path = os.path.dirname(os.path.abspath(sys.argv[0]))
        print(self.exe_dir_path)

        self.width, self.height = width, height
        self.root.geometry("%dx%d-%d+%d" % (width, height, width_adjust, higth_adjust))

        self.root.maxsize(width, 2000)
        self.root.minsize(width, 30)

        # 后下手遭殃，可能会被隐藏
        self.frame_bottom = Frame(self.root, bg=bg)
        self.frame_bottom.pack(side=BOTTOM)
        self.frame_top = Frame(self.root, bg=bg)
        self.frame_top.pack(side=TOP)
        self.frame_middle = Frame(self.root, bg=bg)
        self.frame_middle.pack(side=BOTTOM, padx=21, pady=7)

        ttk.Label(self.frame_top, text="轨迹：", background=bg).pack(side=tk.LEFT, pady=3)
        self.comvalue = tkinter.StringVar()
        self.comboxlist = ttk.Combobox(self.frame_top, textvariable=self.comvalue, width=24, state='readonly')
        self.comboxlist.pack(side=tk.LEFT, pady=3)
        self.refresh_photo_image = ImageTk.PhotoImage(
            Image.open(self.exe_dir_path + "/img/" + "refresh.png").resize((18, 18)))
        tk.Button(self.frame_top,
                  command=self.refresh_files,
                  bg=bg,
                  image=self.refresh_photo_image,
                  relief="flat",
                  overrelief="groove").pack(side=tk.LEFT, padx=5, pady=3)

        # 新建按钮
        self.new_image = Image.open(self.exe_dir_path + "/img/" + "24gl-fileEmpty.png").resize((25, 25))
        self.new_photo_image = ImageTk.PhotoImage(self.new_image)
        self.new_button = tk.Button(self.frame_bottom,
                                    command=self.new_track,
                                    bg=bg,
                                    text="编辑", image=self.new_photo_image,
                                    relief="flat",
                                    overrelief="groove")
        self.new_button.pack(side=tk.LEFT, padx=5, pady=3)

        # 编辑按钮
        self.edit_image = Image.open(self.exe_dir_path + "/img/" + "edit.png").resize((25, 25))
        self.edit_photo_image = ImageTk.PhotoImage(self.edit_image)
        self.edit_button = tk.Button(self.frame_bottom,
                                     command=self.edit_track,
                                     bg=bg,
                                     text="编辑",
                                     image=self.edit_photo_image,
                                     relief="flat",
                                     overrelief="groove")
        self.edit_button.pack(side=tk.LEFT, padx=5, pady=3)

        # 开始按钮
        self.start_photo_image = ImageTk.PhotoImage(Image.open(self.exe_dir_path + "/img/" + "开始.png").resize((25, 25)))
        self.start_button = tk.Button(self.frame_bottom,
                                      command=self.start,
                                      bg=bg,
                                      image=self.start_photo_image,
                                      relief="flat",
                                      overrelief="groove")
        self.start_button.pack(side=tk.LEFT, padx=5, pady=3)

        # 透明按钮
        self.alpha_photo_image = ImageTk.PhotoImage(
            Image.open(self.exe_dir_path + "/img/" + "透明度.png").resize((25, 25)))
        self.alpha_button = tk.Button(self.frame_bottom,
                                      command=self.alpha,
                                      bg=bg,
                                      image=self.alpha_photo_image,
                                      relief="flat",
                                      overrelief="groove")
        self.alpha_button.pack(side=tk.LEFT, padx=5, pady=3)

        # 测试按钮
        self.test_photo_image = ImageTk.PhotoImage(
            Image.open(self.exe_dir_path + "/img/" + "shutter.png").resize((25, 25)))
        self.test_button = tk.Button(self.frame_bottom,
                                     command=self.test_track,
                                     bg=bg,
                                     image=self.test_photo_image,
                                     relief="flat",
                                     overrelief="groove")
        self.test_button.pack(side=tk.LEFT, padx=5, pady=3)

        # 帮助按钮
        self.help_photo_image = ImageTk.PhotoImage(
            Image.open(self.exe_dir_path + "/img/" + "帮助.png").resize((25, 25)))
        self.help_button = tk.Button(self.frame_bottom,
                                     command=self.help,
                                     bg=bg,
                                     image=self.help_photo_image,
                                     relief="flat",
                                     overrelief="groove")
        self.help_button.pack(side=tk.LEFT, padx=5, pady=3)

        # 设置按钮
        self.stop_photo_image = ImageTk.PhotoImage(
            Image.open(self.exe_dir_path + "/img/" + "setting.png").resize((25, 25)))
        self.stop_button = tk.Button(self.frame_bottom,
                                     command=self.setting,
                                     bg=bg,
                                     image=self.stop_photo_image,
                                     relief="flat",
                                     overrelief="groove")
        self.stop_button.pack(side=tk.LEFT, padx=5, pady=3)

        # 开启初始化线程
        self.p = Thread(target=self.init)
        self.p.setDaemon(True)
        self.p.start()

        # 注册监听热键
        keyboard.add_hotkey('ctrl+c', self.listen_track_stop)
        keyboard.add_hotkey('ctrl+z', self.close)

        self.root.mainloop()

    def refresh_files(self):
        self.comboxlist.config(values=os.listdir(self.exe_dir_path + "/track/"))
        self.comboxlist.current(0)

    def test_track(self):
        Test_window(height=200, pre_window=self)

    def setting(self):
        Setting_window(height=160, pre_window=self)

    def close(self):
        print("按住了ctrl+z")
        self.root.destroy()

    def init(self):
        time.sleep(1)
        # 耗时UI放到线程里面
        self.wait_window = Wait_window(height=80, pre_window=self)

        # 是否第一次运行
        if not os.path.exists(self.exe_dir_path + "/zod_setting.sqlite"):
            print("第一次运行")
            my_sql_dict = SqliteDict(self.exe_dir_path + '/zod_setting.sqlite')
            my_sql_dict["coordinate"] = [(500, 300)]
            my_sql_dict["mode"] = ["有蜂鸣器声音"]
            my_sql_dict["title"] = ["Honey"]
            my_sql_dict.commit()
            my_sql_dict.close()

        self.my_dict = SqliteDict(self.exe_dir_path + '/zod_setting.sqlite', autocommit=True)

        # 初始化下拉列表
        self.comboxlist.config(values=os.listdir(self.exe_dir_path + "/track/"))
        self.comboxlist.current(0)
        # 初始化大漠插件
        pythoncom.CoInitialize()
        self.dm = DM(dll_path=os.path.join(self.exe_dir_path,"dm.dll"))
        print(self.dm.GetBasePath())
        # self.dm.Un_reg()

        # 关闭等待窗口
        self.wait_window.close()
        # 监听键盘
        keyboard.wait()

    def listen_track_stop(self):
        print("按住了ctrl+c")
        try:
            self.track.stop()
        except Exception:
            traceback.print_exc()

    def edit_track(self):
        Track_window(height=280, pre_window=self, flag=1)

    def new_track(self):
        Track_window(height=280, pre_window=self, flag=0)

    def start(self):
        self.track = Track(self)
        self.track.start()

    def alpha(self):
        Alpha_window(height=40, pre_window=self)

    def help(self, **kwargs):
        Help_window(width=1000,
                    height=600,
                    path=self.exe_dir_path + "/README.md")


class Track_window(object):
    def __init__(self, height, pre_window, **kwargs):
        self.root = tk.Toplevel()
        self.root.iconbitmap("favicon.ico")
        self.root.wm_attributes('-topmost', 1)

        self.pre_window = pre_window
        self.pre_window_root = pre_window.root
        self.width, self.height = pre_window.width, height
        self.exe_dir_path = pre_window.exe_dir_path
        self.file_name = self.pre_window.comboxlist.get()
        self.flag = kwargs["flag"]

        self.pre_window_root_x, self.pre_window_root_y, self.pre_window_root_w, self.pre_window_root_h = \
            self.pre_window_root.winfo_x(), \
            self.pre_window_root.winfo_y(), \
            self.pre_window_root.winfo_width(), \
            self.pre_window_root.winfo_height()
        self.width_adjust = self.pre_window_root_x
        self.height_adjust = self.pre_window_root_y + self.pre_window_root_h + 40

        self.root.geometry("%dx%d+%d+%d" % (self.width, self.height, self.width_adjust, self.height_adjust))

        # 布局
        self.frame_bottom2 = Frame(self.root)
        self.frame_bottom2.pack(side=BOTTOM, padx=20, pady=5, expand=True, fill=X)
        self.frame_top = Frame(self.root)
        self.frame_top.pack(side=TOP, padx=20, pady=10)
        self.frame_middle = Frame(self.root)
        self.frame_middle.pack(side=TOP, padx=20, pady=5)

        # 轨迹名称
        ttk.Label(self.frame_top, text="轨迹名称").pack(side=tk.LEFT, padx=5, pady=2)
        self.track_name_str = tk.StringVar()
        self.track_name = ttk.Entry(self.frame_top,
                                    validate="focus",
                                    validatecommand=self.clear,
                                    textvariable=self.track_name_str)
        self.track_name.pack(side=tk.LEFT, fill="x", expand=True)
        ttk.Label(self.frame_top, text=".txt").pack(side=tk.LEFT, padx=2, pady=2)
        tk.Label(self.frame_top, state="disable", width=1000).pack(side=tk.LEFT, expand=True, fill=X)

        # 轨迹内容 n代表向上对齐
        ttk.Label(self.frame_middle, text="轨迹内容").pack(anchor="n", side=tk.LEFT, padx=5, pady=2)
        self.track_content = ScrolledText(self.frame_middle)
        self.track_content.pack(side=tk.LEFT)
        tk.Label(self.frame_middle, state="disable", width=1000).pack(side=tk.LEFT, expand=True, fill=X)

        # 完成按钮
        self.save_task_photo_image = ImageTk.PhotoImage(
            Image.open(self.exe_dir_path + "/img/" + "确认.png").resize((25, 25)))
        self.save_task_button = tk.Button(self.frame_bottom2, command=self.ok,
                                          text="完成",
                                          image=self.save_task_photo_image,
                                          relief="flat", overrelief="groove")
        self.save_task_button.pack(side=tk.RIGHT)

        # 删除按钮
        self.del_task_photo_image = ImageTk.PhotoImage(
            Image.open(self.exe_dir_path + "/img/" + "24gl-trash2.png").resize((25, 25)))
        self.del_task_button = tk.Button(self.frame_bottom2, command=self.del_task,
                                         text="完成",
                                         image=self.del_task_photo_image,
                                         relief="flat", overrelief="groove")
        self.del_task_button.pack(side=tk.RIGHT)

        if self.flag == 1:
            self.root.title('修改轨迹')
            self.track_name.insert(0, self.file_name[:-4])
            with open(self.exe_dir_path + "/track/" + self.file_name, "r", encoding="utf8") as f:
                track_str = f.read()
                self.track_content.insert(END, track_str)
        else:
            self.root.title("新建轨迹")
            self.track_name.insert(0, "输入轨迹名称")
            self.del_task_button.pack_forget()

        self.root.mainloop()

    def ok(self):
        if self.flag == 1:
            os.remove(self.exe_dir_path + "/track/" + self.file_name)
        with open(self.exe_dir_path + "/track/" + self.track_name_str.get() + ".txt", "w", encoding="utf-8") as f:
            track_content_get = self.track_content.get(1.0, END)
            f.write(track_content_get)
        self.pre_window.refresh_files()
        self.root.destroy()

    def clear(self):
        if "输入轨迹名称" in self.track_name.get():
            self.track_name.delete(0, "end")

    def del_task(self):
        os.remove(self.exe_dir_path + "/track/" + self.file_name)
        self.pre_window.refresh_files()
        self.root.destroy()


class Alpha_window(object):
    def __init__(self, height, pre_window, **kwargs):
        self.root = tk.Toplevel()
        self.root.title('调整透明度')
        self.root.iconbitmap("favicon.ico")
        self.root.wm_attributes('-topmost', 1)

        self.pre_window = pre_window
        self.pre_window_root = pre_window.root
        self.width = pre_window.width
        self.height = height

        self.pre_window_root_x, self.pre_window_root_y, self.pre_window_root_w, self.pre_window_root_h = \
            self.pre_window_root.winfo_x(), \
            self.pre_window_root.winfo_y(), \
            self.pre_window_root.winfo_width(), \
            self.pre_window_root.winfo_height()
        self.width_adjust = self.pre_window_root_x
        self.height_adjust = self.pre_window_root_y + self.pre_window_root_h + 40

        self.root.geometry("%dx%d+%d+%d" % (self.width, self.height, self.width_adjust, self.height_adjust))

        # 布局
        self.frame_bottom = Frame(self.root)
        self.frame_bottom.pack(side=BOTTOM, fill="x", expand=True, padx=10, pady=5)

        self.scale = ttk.Scale(self.frame_bottom, from_=0, to=100,
                               command=self.set_alpha,
                               value=90)
        self.scale.pack(side=tk.LEFT, padx=5, pady=2, expand=True, fill=X)
        self.scale.set(float(self.pre_window_root.wm_attributes("-alpha")) * 100)
        self.root.mainloop()

    def set_alpha(self, i):
        self.pre_window_root.attributes('-alpha', float(i) / 100)


class Help_window(object):
    def __init__(self, width, height, path):
        self.root = tk.Toplevel()
        self.root.title('使用说明')
        self.root.iconbitmap("favicon.ico")
        self.path = path

        # 居中显示
        self.width, self.height = width, height
        win_width = self.root.winfo_screenwidth()
        win_higth = self.root.winfo_screenheight()
        width_adjust = (win_width - width) / 2
        higth_adjust = (win_higth - height) / 2
        self.root.geometry("%dx%d+%d+%d" % (width, height, width_adjust, higth_adjust))

        self.frame_bottom = ttk.Frame(self.root)
        self.frame_bottom.pack(side=tk.BOTTOM, fill=BOTH, expand=False)
        self.frame_top = ttk.Frame(self.root, padding=20)
        self.frame_top.pack(side=tk.TOP, fill=BOTH, expand=True)

        self.HtmlView = TkHtmlView(self.frame_top, background="white")
        self.HtmlView.pack(fill='both', expand=True)

        self.filename = tk.StringVar()
        ttk.Entry(self.frame_bottom, textvariable=self.filename).pack(side=tk.LEFT, padx=20, pady=10, fill='x',
                                                                      expand=True)
        ttk.Button(self.frame_bottom, text='选择文件', command=self.open_file).pack(side=tk.LEFT, padx=20, pady=10)

        # 加载默认md文件
        self.p = Thread(target=self.init_file)
        self.p.start()

        self.root.mainloop()

    def init_file(self):
        with open(self.path, encoding='utf-8') as f:
            self.filename.set(self.path)
            md2html = Markdown()
            html = md2html.convert(f.read())
            self.HtmlView.set_html(html)

    def open_file(self):
        path = askopenfilename()
        if not path:
            return
        with open(path, encoding='utf-8') as f:
            self.filename.set(path)
            md2html = Markdown()
            html = md2html.convert(f.read())
            print(html)
            self.HtmlView.set_html(html)


class Wait_window(object):
    def __init__(self, height, pre_window):
        self.root = tk.Toplevel()
        self.root.title('初始化')
        self.root.iconbitmap("favicon.ico")
        self.root.wm_attributes('-topmost', 1)

        self.pre_window = pre_window
        self.pre_window_root = pre_window.root
        self.width = pre_window.width
        self.height = height

        self.pre_window_root_x, self.pre_window_root_y, self.pre_window_root_w, self.pre_window_root_h = \
            self.pre_window_root.winfo_x(), \
            self.pre_window_root.winfo_y(), \
            self.pre_window_root.winfo_width(), \
            self.pre_window_root.winfo_height()
        self.width_adjust = self.pre_window_root_x
        self.height_adjust = self.pre_window_root_y + self.pre_window_root_h + 40

        self.root.geometry("%dx%d+%d+%d" % (self.width, self.height, self.width_adjust, self.height_adjust))

        # 布局
        self.frame_bottom = Frame(self.root)
        self.frame_bottom.pack(side=BOTTOM, padx=15, pady=10, expand=True, fill=X)
        self.frame_top = Frame(self.root)
        self.frame_top.pack(side=TOP, padx=5, pady=10)

        # 进度条
        self.bar = ttk.Progressbar(self.frame_bottom, mode="indeterminate", orient=tk.HORIZONTAL)
        self.bar.pack(expand=True, fill=X)
        self.bar.start(10)

        # 提示内容
        self.content = tk.Label(self.frame_top, text="正在初始化,请不要操作，请耐心等待......")
        self.content.pack()

        # self.root.mainloop()

    def close(self):
        self.root.destroy()


class Setting_window(object):
    def __init__(self, height, pre_window, **kwargs):
        self.root = tk.Toplevel()
        self.root.title('设置')
        self.root.iconbitmap("favicon.ico")
        self.root.wm_attributes('-topmost', 1)

        self.exe_dir_path = pre_window.exe_dir_path

        self.pre_window = pre_window
        self.pre_window_root = pre_window.root
        self.width = pre_window.width
        self.height = height

        self.pre_window_root_x, self.pre_window_root_y, self.pre_window_root_w, self.pre_window_root_h = \
            self.pre_window_root.winfo_x(), \
            self.pre_window_root.winfo_y(), \
            self.pre_window_root.winfo_width(), \
            self.pre_window_root.winfo_height()
        self.width_adjust = self.pre_window_root_x
        self.height_adjust = self.pre_window_root_y + self.pre_window_root_h + 40
        self.root.geometry("%dx%d+%d+%d" % (self.width, self.height, self.width_adjust, self.height_adjust))

        self.my_dict = pre_window.my_dict
        self.coordinate_x = tk.StringVar()
        self.coordinate_y = tk.StringVar()
        self.mode_combobox_str = tk.StringVar()
        self.title_str = tk.StringVar()

        self.coordinate = self.my_dict["coordinate"][0]
        self.coordinate_x.set(self.coordinate[0])
        self.coordinate_y.set(self.coordinate[1])
        self.mode_combobox_str.set(self.my_dict["mode"][0])
        self.title_str.set(self.my_dict["title"][0])

        # 布局
        self.frame_top = Frame(self.root)
        self.frame_top.pack(side=TOP, padx=20, pady=10)

        self.frame_middle2 = Frame(self.root)
        self.frame_middle2.pack(side=TOP, padx=20, pady=1)

        self.frame_middle = Frame(self.root)
        self.frame_middle.pack(side=TOP, padx=20, pady=10)

        self.frame_bottom = Frame(self.root)
        self.frame_bottom.pack(side=BOTTOM, padx=20, expand=True, fill=X)

        # 基准
        ttk.Label(self.frame_top, text="基准坐标：").pack(side=tk.LEFT, padx=5, pady=2)
        self.coordinate_x_entry = ttk.Entry(self.frame_top, width=5, textvariable=self.coordinate_x)
        self.coordinate_x_entry.pack(side=tk.LEFT, padx=3)
        self.coordinate_y_entry = ttk.Entry(self.frame_top, width=5, textvariable=self.coordinate_y)
        self.coordinate_y_entry.pack(side=tk.LEFT, padx=3)
        tk.Label(self.frame_top, state="disable", width=1000).pack(side=tk.LEFT, expand=True, fill=X)

        # 模式
        ttk.Label(self.frame_middle, text="启动模式：").pack(side=tk.LEFT, padx=5, pady=2)
        self.mode_combobox = ttk.Combobox(self.frame_middle, width=26, values=["有蜂鸣器声音", "没有蜂鸣器声音"], state="readonly",
                                          textvariable=self.mode_combobox_str)
        self.mode_combobox.pack(side=tk.LEFT, padx=0, pady=2)
        tk.Label(self.frame_middle, state="disable", width=1000).pack(side=tk.LEFT, expand=True, fill=X)

        # 标题
        ttk.Label(self.frame_middle2, text="窗口标题：").pack(side=tk.LEFT, padx=5, pady=2)
        self.title_entry = ttk.Entry(self.frame_middle2, width=10, textvariable=self.title_str)
        self.title_entry.pack(side=tk.LEFT, padx=3)
        tk.Label(self.frame_middle2, text="模糊匹配", width=1000).pack(side=tk.LEFT, expand=True, fill=X)

        # 按钮
        self.cancel_image = ImageTk.PhotoImage(Image.open(self.exe_dir_path + "/img/" + "取消.png").resize((18, 18)))
        self.cancel_button = tk.Button(self.frame_bottom,
                                       command=self.cancel,
                                       image=self.cancel_image,
                                       relief="flat",
                                       overrelief="groove")
        self.cancel_button.pack(side=tk.RIGHT, padx=5, pady=2)

        self.ok_image = ImageTk.PhotoImage(Image.open(self.exe_dir_path + "/img/" + "确认.png").resize((18, 18)))
        self.ok_button = tk.Button(self.frame_bottom,
                                   command=self.ok,
                                   image=self.ok_image,
                                   relief="flat",
                                   overrelief="groove")
        self.ok_button.pack(side=tk.RIGHT, padx=5, pady=2)

        self.root.mainloop()

    def cancel(self):
        self.root.destroy()

    def ok(self):
        self.my_dict["coordinate"] = [(eval(self.coordinate_x.get()), eval(self.coordinate_y.get()))]
        self.my_dict["mode"] = [self.mode_combobox_str.get()]
        self.my_dict["title"] = [self.title_str.get()]

        self.root.destroy()


class Track(threading.Thread):
    def __init__(self, pre_window, *args, **kwargs):
        super(Track, self).__init__(*args, **kwargs)
        self.__flag = threading.Event()  # 用于暂停线程的标识
        self.__flag.set()  # 设置为True
        self.__running = threading.Event()  # 用于停止线程的标识
        self.__running.set()  # 将running设置为True

        self.pre_window = pre_window
        self.dm = pre_window.dm
        self.exe_dir_path = pre_window.exe_dir_path
        self.my_dict = pre_window.my_dict

        self.file_name = self.pre_window.comboxlist.get()
        self.command_dict_list = []
        self.read_file(file_name=self.exe_dir_path + "/track/" + self.file_name)

    def read_file(self, file_name):
        self.step_dict_list = []
        with open(file_name, "r", encoding="utf8") as f:
            for line in f.readlines():
                linestr = line.strip()
                if linestr != "":
                    linestrlist = linestr.split("-")
                    self.step_dict_list.append({"command": linestrlist[0], "content": linestrlist})

    def run(self):
        if self.ready() == 0:
            return
        for step_dict in self.step_dict_list:
            self.__flag.wait()  # 为True时立即返回, 为False时阻塞直到内部的标识位为True后返回
            if self.__running.isSet():

                self.step_dict = step_dict
                print(step_dict)
                StepCommand(self.dm, self.hwnd, self.coordinate, self.step_dict)

            else:
                break

    def ready(self):
        """
        操作前的准备
        :return:
        """
        self.coordinate = self.my_dict["coordinate"][0]
        self.title = self.my_dict["title"][0]
        if self.my_dict["mode"][0] == "有蜂鸣器声音":
            self.dm.Beep()

        # 找到窗体
        self.hwnd = self.dm.FindWindow("", self.title)
        if self.hwnd == 0:
            tkinter.messagebox.askokcancel(title='没有找到窗口',
                                           message='确定你的游戏是否运行？或者你的窗口标题是否正确？')
            return 0
        else:
            # 激活窗体
            self.dm.SetWindowState(self.hwnd, 1)
            return 1

    def stop(self):
        self.__flag.set()  # 将线程从暂停状态恢复, 如何已经暂停的话
        self.__running.clear()  # 设置为False


class Test_window(object):
    def __init__(self, height, pre_window, **kwargs):
        self.root = tk.Toplevel()
        self.root.title('测试轨迹')
        self.root.iconbitmap("favicon.ico")
        self.root.wm_attributes('-topmost', 1)

        self.my_dict = pre_window.my_dict
        self.pre_window = pre_window
        self.pre_window_root = pre_window.root
        self.exe_dir_path = pre_window.exe_dir_path
        self.width = pre_window.width
        self.height = height
        self.dm = pre_window.dm

        self.pre_window_root_x, self.pre_window_root_y, self.pre_window_root_w, self.pre_window_root_h = \
            self.pre_window_root.winfo_x(), \
            self.pre_window_root.winfo_y(), \
            self.pre_window_root.winfo_width(), \
            self.pre_window_root.winfo_height()
        self.width_adjust = self.pre_window_root_x
        self.height_adjust = self.pre_window_root_y + self.pre_window_root_h + 40

        self.root.geometry("%dx%d+%d+%d" % (self.width, self.height, self.width_adjust, self.height_adjust))

        # 布局
        self.frame_top = Frame(self.root)
        self.frame_top.pack(side=TOP, pady=5, padx=15)
        self.frame_middle = Frame(self.root)
        self.frame_middle.pack(side=TOP, pady=5, padx=15)
        self.frame_middle2 = Frame(self.root)
        self.frame_middle2.pack(side=TOP, pady=5, padx=15)
        self.frame_middle3 = Frame(self.root)
        self.frame_middle3.pack(side=TOP, pady=5, padx=15)
        self.frame_bottom = Frame(self.root)
        self.frame_bottom.pack(side=BOTTOM, padx=20, pady=5, expand=True, fill=X)

        # 只能右侧，左侧用框架控制
        ttk.Label(self.frame_top, text="指令：").pack(side=tk.LEFT, pady=3)
        self.command_comvalue = tkinter.StringVar()
        self.comboxlist = ttk.Combobox(self.frame_top, textvariable=self.command_comvalue, width=27)
        self.comboxlist["values"] = ("请选择指令")
        self.comboxlist['state'] = 'readonly'
        self.comboxlist.current(0)
        self.comboxlist.pack(side=tk.LEFT, pady=3)
        tk.Label(self.frame_top, state="disable", width=1000).pack(side=tk.LEFT, expand=True, fill=X)
        self.comboxlist.bind("<<ComboboxSelected>>", self.manage)

        ttk.Label(self.frame_middle, text="步长：").pack(side=tk.LEFT, pady=3)
        self.step_value = tkinter.StringVar()
        self.step_entry = ttk.Entry(self.frame_middle, textvariable=self.step_value, width=20)
        self.step_entry.pack(side=tk.LEFT, pady=3)
        tk.Label(self.frame_middle, text="单位是像素").pack(side=tk.LEFT)
        tk.Label(self.frame_middle, state="disable", width=1000).pack(side=tk.LEFT, expand=True, fill=X)

        ttk.Label(self.frame_middle2, text="时长：").pack(side=tk.LEFT, pady=3)
        self.duration_value = tkinter.StringVar()
        self.duration_entry = ttk.Entry(self.frame_middle2, textvariable=self.duration_value, width=20)
        self.duration_entry.pack(side=tk.LEFT, pady=3)
        tk.Label(self.frame_middle2, text="单位是毫秒").pack(side=tk.LEFT)
        tk.Label(self.frame_middle2, state="disable", width=1000).pack(side=tk.LEFT, expand=True, fill=X)

        ttk.Label(self.frame_middle3, text="生成：").pack(side=tk.LEFT, pady=3)
        self.output_value = tkinter.StringVar()
        self.output_entry = ttk.Entry(self.frame_middle3, textvariable=self.output_value, width=27)
        self.output_entry.pack(side=tk.LEFT, pady=3)
        tk.Label(self.frame_middle3, state="disable", width=1000).pack(side=tk.LEFT, expand=True, fill=X)

        # 确定按钮
        self.do_photo_image = ImageTk.PhotoImage(
            Image.open(self.exe_dir_path + "/img/" + "确认.png").resize((25, 25)))
        self.do_button = tk.Button(self.frame_bottom,
                                   command=self.work,
                                   text="确定",
                                   image=self.do_photo_image,
                                   relief="flat", overrelief="groove")
        self.do_button.pack(side=tk.RIGHT)

        # 撤回按钮
        self.undo_photo_image = ImageTk.PhotoImage(
            Image.open(self.exe_dir_path + "/img/" + "undo.png").resize((25, 25)))
        self.undo_button = tk.Button(self.frame_bottom,
                                     command=self.undo,
                                     text="撤回",
                                     image=self.undo_photo_image,
                                     relief="flat", overrelief="groove")
        self.undo_button.pack(side=tk.RIGHT)

        self.key_single_list = ["复位相机", "朝向脸部", "朝向胸部", "朝向跨下", "加速抽送", "减速抽送"]
        self.key_press_list = ["向前平移", "向后平移", "向左平移", "向右平移", "向上平移", "向下平移", "旋转视角"]
        self.left_click_list = ["向上拖动", "向下拖动", "向左拖动", "向右拖动", "左上拖动", "左下拖动", "右上拖动", "右下拖动"]
        # self.custom_list = ["自定义左键", "自定义右键"]
        self.custom_list = []
        self.all_list = self.key_single_list + self.key_press_list + self.left_click_list + self.custom_list
        self.comboxlist.config(value=self.all_list)

        self.root.mainloop()

    def manage(self, event):
        self.command_comvalue_str = self.command_comvalue.get()
        if self.command_comvalue_str in self.key_single_list:
            self.step_entry.config(state=tk.DISABLED)
            self.duration_entry.config(state=tk.DISABLED)
            self.undo_button.pack_forget()

        elif self.command_comvalue_str in self.key_press_list:
            self.step_entry.config(state=tk.DISABLED)
            self.duration_entry.config(state=tk.NORMAL)
            self.undo_button.pack(side=tk.RIGHT)

        elif self.command_comvalue_str in self.left_click_list:
            self.step_entry.config(state=tk.NORMAL)
            self.duration_entry.config(state=tk.NORMAL)
            self.undo_button.pack(side=tk.RIGHT)

    def undo(self):
        self.output_value_str = self.output_value.get()

        self.output_value_str_ = ""
        for o in self.output_value_str:
            if o == "上":
                self.output_value_str_ = self.output_value_str_ + "下"
            elif o == "下":
                self.output_value_str_ = self.output_value_str_ + "上"
            elif o == "左":
                self.output_value_str_ = self.output_value_str_ + "右"
            elif o == "右":
                self.output_value_str_ = self.output_value_str_ + "左"
            elif o == "前":
                self.output_value_str_ = self.output_value_str_ + "后"
            elif o == "后":
                self.output_value_str_ = self.output_value_str_ + "前"
            else:
                self.output_value_str_ = self.output_value_str_ + o

        self.begin(self.output_value_str_.split("-"))

    def work(self):
        # 生成指令
        self.command_comvalue_str = self.command_comvalue.get()
        if self.command_comvalue_str in self.key_single_list:
            self.command_comvalue_str = self.command_comvalue.get()
            self.output_value.set(self.command_comvalue_str)
        elif self.command_comvalue_str in self.key_press_list:
            self.command_comvalue_str = self.command_comvalue.get() + "-" + self.duration_value.get()
            self.output_value.set(self.command_comvalue_str)
        elif self.command_comvalue_str in self.left_click_list:
            self.command_comvalue_str = self.command_comvalue.get() + "-" + self.step_value.get() + "-" + self.duration_value.get()
            self.output_value.set(self.command_comvalue_str)

        self.begin(self.command_comvalue_str.split("-"))

    def begin(self, linestrlist):
        if self.ready() == 0:
            return
        self.step_dict = {"command": linestrlist[0], "content": linestrlist}
        print(self.step_dict)
        StepCommand(self.dm, self.hwnd, self.coordinate, self.step_dict)

    def ready(self):
        """
        操作前的准备
        :return:
        """
        self.coordinate = self.my_dict["coordinate"][0]
        self.title = self.my_dict["title"][0]
        if self.my_dict["mode"][0] == "有蜂鸣器声音":
            self.dm.Beep()

        # 找到窗体
        self.hwnd = self.dm.FindWindow("", self.title)
        if self.hwnd == 0:
            tkinter.messagebox.askokcancel(title='没有找到窗口',
                                           message='确定你的游戏是否运行？或者你的窗口标题是否正确？')
            return 0
        else:
            # 激活窗体
            self.dm.SetWindowState(self.hwnd, 1)
            return 1


class StepCommand:
    '''
    向指定窗体发送命令
    '''

    def __init__(self, dm, hwnd, coordinate, step_dict):
        self.dm = dm
        self.hwnd = hwnd

        (self.center_x, self.center_y) = coordinate
        self.dm.MoveTo(self.center_x, self.center_y)

        flag = self.dm.GetWindowState(self.hwnd, 1)
        if not flag:
            self.dm.SetWindowState(self.hwnd, 1)
        # 按键
        if step_dict["command"] == "复位相机":
            self.cam_init(step_dict)
        elif step_dict["command"] == "朝向脸部":
            self.cam_to_face(step_dict)
        elif step_dict["command"] == "朝向胸部":
            self.cam_to_chest(step_dict)
        elif step_dict["command"] == "朝向跨下":
            self.cam_to_crawl(step_dict)
        elif step_dict["command"] == "加速抽送":
            self.speed_up(step_dict)
        elif step_dict["command"] == "减速抽送":
            self.speed_down(step_dict)

        # 按键-时长
        elif step_dict["command"] == "向前平移":
            self.translation_front(step_dict)
        elif step_dict["command"] == "向后平移":
            self.translation_back(step_dict)
        elif step_dict["command"] == "向左平移":
            self.translation_left(step_dict)
        elif step_dict["command"] == "向右平移":
            self.translation_right(step_dict)
        elif step_dict["command"] == "向上平移":
            self.translation_up(step_dict)
        elif step_dict["command"] == "向下平移":
            self.translation_down(step_dict)
        elif step_dict["command"] == "旋转视角":
            self.rotate(step_dict)

        # 左击-拖动-时长
        elif step_dict["command"] == "向上拖动":
            self.turn_up(step_dict)
        elif step_dict["command"] == "向下拖动":
            self.turn_down(step_dict)
        elif step_dict["command"] == "向左拖动":
            self.turn_left(step_dict)
        elif step_dict["command"] == "向右拖动":
            self.turn_right(step_dict)
        elif step_dict["command"] == "左上拖动":
            self.turn_left_up(step_dict)
        elif step_dict["command"] == "左下拖动":
            self.turn_left_down(step_dict)
        elif step_dict["command"] == "右上拖动":
            self.turn_right_up(step_dict)
        elif step_dict["command"] == "右下拖动":
            self.turn_right_down(step_dict)

        # 自定义
        elif step_dict["command"] == "自定义左键":
            self.custom_left(step_dict)
        elif step_dict["command"] == "自定义右键":
            self.custom_right(step_dict)

    # 单纯按键类
    # 复位相机
    def cam_init(self, step_dict):
        self.dm.KeyPressChar("r")

    # 朝向脸部
    def cam_to_face(self, step_dict):
        self.dm.KeyPressChar("q")

    # 朝向胸部
    def cam_to_chest(self, step_dict):
        self.dm.KeyPressChar("w")

    # 朝向跨下
    def cam_to_crawl(self, step_dict):
        self.dm.KeyPressChar("e")

    # 加速抽送
    def speed_up(self, step_dict):
        self.dm.WheelUp()

    # 减速抽送
    def speed_down(self, step_dict):
        self.dm.WheelDown()

    # 按键时间类
    # 向后平移
    def translation_back(self, step_dict):
        duration = eval(step_dict["content"][1])
        s_duration = round(duration / 1000, 4)
        self.dm.KeyDownChar("down")
        time.sleep(s_duration)
        self.dm.KeyUpChar("down")

    # 向前平移
    def translation_front(self, step_dict):
        duration = eval(step_dict["content"][1])
        s_duration = round(duration / 1000, 4)
        self.dm.KeyDownChar("up")
        time.sleep(s_duration)
        self.dm.KeyUpChar("up")

    # 向左平移
    def translation_left(self, step_dict):
        duration = eval(step_dict["content"][1])
        s_duration = round(duration / 1000, 4)
        self.dm.KeyDownChar("left")
        time.sleep(s_duration)
        self.dm.KeyUpChar("left")

    # 向右平移
    def translation_right(self, step_dict):
        duration = eval(step_dict["content"][1])
        s_duration = round(duration / 1000, 4)
        self.dm.KeyDownChar("right")
        time.sleep(s_duration)
        self.dm.KeyUpChar("right")

    # 向上平移
    def translation_up(self, step_dict):
        duration = eval(step_dict["content"][1])
        s_duration = round(duration / 1000, 4)
        self.dm.KeyDownChar("pgup")
        time.sleep(s_duration)
        self.dm.KeyUpChar("pgup")

    # 向下平移
    def translation_down(self, step_dict):
        duration = eval(step_dict["content"][1])
        s_duration = round(duration / 1000, 4)
        self.dm.KeyDownChar("pgdn")
        time.sleep(s_duration)
        self.dm.KeyUpChar("pgdn")

    # 旋转视角
    def rotate(self, step_dict):
        duration = eval(step_dict["content"][1])
        s_duration = round(duration / 1000, 4)
        self.dm.KeyDownChar("\\")
        time.sleep(s_duration)
        self.dm.KeyUpChar("\\")

    # 鼠标左类
    # 向下拖动
    def turn_down(self, step_dict):
        length = eval(step_dict["content"][1])
        duration = eval(step_dict["content"][2])
        self.dm.LeftDown()
        pre_duration = round(duration / 1000 / length, 4)
        for i in range(length):
            time.sleep(pre_duration)
            self.dm.MoveR(0, 1)
        self.dm.LeftUp()

    # 向上拖动
    def turn_up(self, step_dict):
        length = eval(step_dict["content"][1])
        duration = eval(step_dict["content"][2])
        self.dm.LeftDown()
        pre_duration = round(duration / 1000 / length, 4)
        for i in range(length):
            time.sleep(pre_duration)
            self.dm.MoveR(0, -1)
        self.dm.LeftUp()

    # 向左拖动
    def turn_left(self, step_dict):
        length = eval(step_dict["content"][1])
        duration = eval(step_dict["content"][2])
        self.dm.LeftDown()
        pre_duration = round(duration / 1000 / length, 4)
        for i in range(length):
            time.sleep(pre_duration)
            self.dm.MoveR(-1, 0)
        self.dm.LeftUp()

    # 向右拖动
    def turn_right(self, step_dict):
        length = eval(step_dict["content"][1])
        duration = eval(step_dict["content"][2])
        self.dm.LeftDown()
        pre_duration = round(duration / 1000 / length, 4)
        for i in range(length):
            time.sleep(pre_duration)
            self.dm.MoveR(1, 0)
        self.dm.LeftUp()

    # 左上拖动
    def turn_left_up(self, step_dict):
        length = eval(step_dict["content"][1])
        duration = eval(step_dict["content"][2])
        self.dm.LeftDown()
        pre_duration = round(duration / 1000 / length, 4)
        for i in range(length):
            time.sleep(pre_duration)
            self.dm.MoveR(-1, -1)
        self.dm.LeftUp()

    # 左下拖动
    def turn_left_down(self, step_dict):
        length = eval(step_dict["content"][1])
        duration = eval(step_dict["content"][2])
        self.dm.LeftDown()
        pre_duration = round(duration / 1000 / length, 4)
        for i in range(length):
            time.sleep(pre_duration)
            self.dm.MoveR(-1, 1)
        self.dm.LeftUp()

    # 右上拖动
    def turn_right_up(self, step_dict):
        length = eval(step_dict["content"][1])
        duration = eval(step_dict["content"][2])
        self.dm.LeftDown()
        pre_duration = round(duration / 1000 / length, 4)
        for i in range(length):
            time.sleep(pre_duration)
            self.dm.MoveR(1, -1)
        self.dm.LeftUp()

    # 右下拖动
    def turn_right_down(self, step_dict):
        length = eval(step_dict["content"][1])
        duration = eval(step_dict["content"][2])
        self.dm.LeftDown()
        pre_duration = round(duration / 1000 / length, 4)
        for i in range(length):
            time.sleep(pre_duration)
            self.dm.MoveR(1, 1)
        self.dm.LeftUp()

    # 自定义左键
    def custom_left(self, step_dict):
        pass

    # 自定义右键
    def custom_right(self, step_dict):
        pass


if __name__ == "__main__":
    main = Main_window(title="ZOD小导演",
                       icon="favicon.ico",
                       alpha=0.9,
                       topmost=1,
                       bg="white",
                       width=300,
                       height=70,
                       width_adjust=10,
                       higth_adjust=5)

# pyinstaller -F -i "C:\Users\Gao yongxian\PycharmProjects\Zod_director\pack\favicon.ico" "C:\Users\Gao yongxian\PycharmProjects\Zod_director\zod.py" -w
