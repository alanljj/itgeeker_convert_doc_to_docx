# -*- coding: utf-8 -*-
###########################################################################
#    Copyright 2023 奇客罗方智能科技 https://www.geekercloud.com
#    ITGeeker.net <alanljj@gmail.com>
############################################################################
import base64
import glob
import json
import os
import tempfile
import tkinter as tk
from tkinter import ttk
import tkinter.filedialog
import tkinter.messagebox
from webbrowser import open_new_tab

from convert_doc_to_docx_api import convert_doc2docx_by_win32com


class AppConvertDoc(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)
        self.label_file_nmb = None
        self.checkbutton = None
        self.include_sub_dir = None
        self.treeview = None
        self.mode_switch = None
        self.entry_path = None
        self.manipulate_frame()
        self.list_frame()
        self.author_frame()

    def select_multiple(self, a):
        cur_items = self.treeview.selection()
        tk.Label(geekerWin, text="\n".join([str(self.treeview.item(i)['values']) for i in cur_items])).pack()

    def cur_select_item(self, a):
        cur_item = self.treeview.focus()
        print(self.treeview.item(cur_item))

    def select_children(self, item):
        # Make sure the item is expanded so the user can see it.
        self.treeview.item(item, open=True)
        # Select the current item.
        self.treeview.selection_add(item)
        # Select all the children of the current item, if any.
        item_children = self.treeview.get_children(item)
        if item_children:
            for sub_item in item_children:
                self.select_children(sub_item)

    def select_all(self):
        for item in self.treeview.get_children():
            self.select_children(item)

    def select_none(self):
        for item in self.treeview.selection():
            self.treeview.selection_remove(item)

    def select_remove(self):
        for item in self.treeview.selection():
            self.treeview.delete(item)

    def delete_items(self, _):
        print('delete')
        for i in self.treeview.selection():
            self.treeview.delete(i)

    def check_sub_dir(self):
        # print('checkbutton.state: ', self.checkbutton.state(["!selected"])) # checkbutton.state:  ('selected',)
        if not self.entry_path.get() or self.entry_path.get() == '浏览并选择目录':
            self.popup_message('no_entry_path')
        else:
            self.list_all_doc_to_tree_view(self.entry_path.get())

    def list_all_doc_to_tree_view(self, root_path):
        files = []
        doc_l = glob.iglob(root_path + r"\[!~$]*.doc")
        if self.include_sub_dir.get():
            doc_l = glob.iglob(root_path + r"\**\[!~$]*.doc", recursive=True)
        print('doc_l: %s' % doc_l)
        if doc_l:
            for docf in doc_l:
                print(docf)
                basename = os.path.basename(docf)
                dirname = os.path.dirname(docf)
                files.append(tuple([basename, dirname]))
        self.treeview.delete(*self.treeview.get_children())
        if files:
            item_count = len(files)
            print('item_count: ', item_count)
            self.label_file_nmb.config(text='文件数：' + str(item_count))
            # self.treeview.delete(*self.treeview.get_children())
            for file_tuple in files:
                self.treeview.insert('', tk.END, values=file_tuple)

    def generate_json_ffp(self):
        cur_usr_path = os.environ['USERPROFILE']
        print('cur_usr_path: %s' % cur_usr_path)
        json_ffp = os.path.join(cur_usr_path, 'itgeeker_convert_doc_to_docx.json')
        if not os.path.isfile(json_ffp):
            ffp_d = dict()
            with open(json_ffp, 'w', encoding='utf-8') as fp:
                fp.write(json.dumps(ffp_d, indent=4, ensure_ascii=False))
                # pass
            return False
        return json_ffp

    def get_all_item_list(self):
        selected_values = []
        # selected_values = self.treeview.focus()
        selected_items = self.treeview.selection()
        if selected_items:
            for sitem in selected_items:
                svalue = self.treeview.item(sitem)
                print('svalue: %s' % type(svalue))  # <class 'dict'>
                print('svalue.details: %s' % svalue.get('values'))
                selected_values.append(svalue.get('values'))
        print('selected_values: %s' % selected_values)
        return list(selected_values)

    def popup_message(self, msg_type):
        if msg_type == 'no_entry_path':
            tk.messagebox.showwarning(title="操作提醒", message="请先选择文件的目录！")

    def start_convert_process(self):
        if not self.entry_path.get() or self.entry_path.get() == '浏览并选择目录':
            self.popup_message('no_entry_path')
        else:
            val_list = self.get_all_item_list()
            if val_list:
                self.save_all_item_to_json(val_list)
                success_l, failed_l = convert_doc2docx_by_win32com(val_list)
                success_nmb = len(success_l)
                failed_nmb = len(failed_l)
                msg_str = "任务已圆满完成！成功处理了%s个文件，%s个文件处理失败\n" \
                          "转换后的.docx文件保存在源文件相同目录，文件名带有-converted字样。" % (str(success_nmb), str(failed_nmb))
                if failed_l:
                    msg_str += "\n本次失败的文件：\n%s" % '\n'.join(x for x in failed_l)
                tk.messagebox.showinfo(title="任务通知", message=msg_str)
                # tk.messagebox.showinfo(title="任务通知",
                #                        message="任务已圆满完成！成功处理了%s个文件，%s个文件处理失败\n"
                #                                "转换后的.docx文件保存在源文件相同目录，文件名带有-converted字样。"
                #                                % (str(success_nmb), str(failed_nmb)))
                if failed_l:
                    cur_usr_path = os.environ['USERPROFILE']
                    failed_ffp = os.path.join(cur_usr_path, 'itgeeker_convert_doc_failed_files.json')
                    with open(failed_ffp, 'w', encoding='utf-8') as ff:
                        failed_d = dict()
                        failed_d['convert_failed_files'] = failed_l
                        ff.write(json.dumps(failed_d, indent=4, ensure_ascii=False))
            else:
                tk.messagebox.showwarning(title="操作提醒", message="请选择要转换的文件，可按住Ctrl多选！")

    def save_all_item_to_json(self, value_list):
        print("here should to save all")
        ffp_d = dict()
        json_ffp = self.generate_json_ffp()

        # print('file_dir: %s' % self.entry_path.get())
        if self.entry_path.get():
            ffp_d['entry_path'] = self.entry_path.get()

        ffp_d['include_sub_dir'] = False
        if self.include_sub_dir.get():
            print('self.include_sub_dir.get()@save: ', self.include_sub_dir.get())
            ffp_d.update({
                'include_sub_dir': True
            })

        ffp_d['label_file_nmb'] = False
        lfn_str = self.label_file_nmb.cget("text")
        print('lfn_str: ', lfn_str)
        if '：' in lfn_str:
            fnmb = lfn_str.split('：')[1]
            ffp_d['label_file_nmb'] = int(fnmb)

        print('ffp_d: ', ffp_d)
        with open(json_ffp, 'w', encoding='utf-8') as ffp:
            file_list = []
            for val_l in value_list:
                f_name = val_l[0]
                f_dir = val_l[1]
                f_dict = {
                    '文件名': f_name,
                    '目录': f_dir
                }
                file_list.append(f_dict)
            ffp_d['file_list'] = file_list
            ffp.write(json.dumps(ffp_d, indent=4, ensure_ascii=False))

    """
    {
        "file_list": [
            {"文件名": "test file name", "目录": "D:\\test"},
            {"文件名": "test file 2", "目录": "D:\\test\\2"},
            {"文件名": "test file 3", "目录": "D:\\test\\2"},
            {"文件名": "test file 3", "目录": "D:\\test\\2"},
            {"文件名": "test file 3", "目录": "D:\\test\\2"},
            {"文件名": "test file 3", "目录": "D:\\test\\2"},
            {"文件名": "test file 3", "目录": "D:\\test\\2"},
            {"文件名": "test file 3", "目录": "D:\\test\\2"},
            {"文件名": "test file 3", "目录": "D:\\test\\2"}
        ],
        "entry_path": "D:\\test\\path"
    }
    """

    def read_all_item_to_treeview_list(self):
        json_ffp = self.generate_json_ffp()
        if json_ffp:
            with open(json_ffp, 'r', encoding='utf-8') as ffp:
                dt_dict = json.load(ffp)
                print('dt_dict: %s' % dt_dict)
            # # keys = tuple(dt.keys())
            # keys = ('文件名', '目录')
            # for col_name in keys:
            #     self.treeview.heading(col_name, text=col_name)
            if 'file_list' in dt_dict:
                self.treeview.delete(*self.treeview.get_children())
                for dt in dt_dict['file_list']:
                    value_tuple = tuple(dt.values())
                    self.treeview.insert('', tk.END, values=value_tuple)
            if 'entry_path' in dt_dict:
                self.entry_path.delete(0, tk.END)
                self.entry_path.insert(0, dt_dict['entry_path'])
            if 'include_sub_dir' in dt_dict:
                print('include_sub_dir: ', dt_dict['include_sub_dir'])
                self.include_sub_dir.set(dt_dict['include_sub_dir'])
            if 'label_file_nmb' in dt_dict:
                # print('label_file_nmb: ', dt_dict['label_file_nmb'])
                if dt_dict['label_file_nmb']:
                    self.label_file_nmb.config(text='文件数：' + str(dt_dict['label_file_nmb']))

    def select_directory(self):
        directory = tk.filedialog.askdirectory()
        self.entry_path.delete(0, tk.END)
        self.entry_path.insert(0, directory)
        self.list_all_doc_to_tree_view(directory)

    def toggle_mode(self):
        if self.mode_switch.instate(["selected"]):
            style.theme_use("forest-light")
        else:
            style.theme_use("forest-dark")

    def open_website(self, url):
        open_new_tab(url)

    def on_window_close(self):
        print("Window closed")
        val_list_on_close = []
        for child in self.treeview.get_children():
            # print(self.treeview.item(child)["values"])
            val_list_on_close.append(self.treeview.item(child)["values"])
        # if val_list_on_close:
        self.save_all_item_to_json(val_list_on_close)
        geekerWin.destroy()

    def manipulate_frame(self):
        mnplt_frame = ttk.LabelFrame(self, text="待转换文件目录")
        mnplt_frame.grid(row=0, column=0, columnspan=3, padx=10, pady=10, ipadx=10, sticky='nsew')

        # path of doc files
        self.entry_path = ttk.Entry(mnplt_frame, justify=tk.LEFT, width=80,
                                    font=('Microsoft YaHei UI', 11))
        self.entry_path.insert(0, "浏览并选择目录")
        # self.entry_path.bind("<FocusIn>", lambda e: self.entry_path.delete('0', 'end'))
        # self.entry_path.focus_force()
        self.entry_path.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

        browse_button = ttk.Button(mnplt_frame, text="选择目录", command=self.select_directory)
        browse_button.grid(row=0, column=2, padx=10, pady=10, ipadx=10, ipady=5, sticky="e")

    def list_frame(self):
        # list_up_frame = ttk.LabelFrame(self, text="选择")
        list_up_frame = ttk.Frame(self)
        list_up_frame.grid(row=1, column=0, columnspan=3, padx=20, pady=10)

        up_sub_fram = ttk.Frame(list_up_frame)
        # up_sub_fram = ttk.LabelFrame(list_up_frame, text="Group")
        up_sub_fram.grid(row=0, column=0, padx=10, pady=5)

        select_all_btn = tk.Button(up_sub_fram, text="选择全部", command=self.select_all,
                                   font=('Microsoft YaHei UI', 11, 'normal'))
        select_all_btn.grid(row=0, column=0, padx=25, pady=5, ipadx=12, ipady=3, sticky="w")

        select_none_btn = tk.Button(up_sub_fram, text="取消选择", command=self.select_none,
                                    font=('Microsoft YaHei UI', 11, 'normal'))
        select_none_btn.grid(row=0, column=1, padx=25, pady=5, ipadx=12, ipady=3, sticky="w")

        select_remove_btn = tk.Button(up_sub_fram, text="移除所选", command=self.select_remove,
                                      font=('Microsoft YaHei UI', 11, 'normal'))
        select_remove_btn.grid(row=0, column=2, padx=25, pady=5, ipadx=12, ipady=3, sticky="w")

        self.include_sub_dir = tk.BooleanVar()
        self.checkbutton = ttk.Checkbutton(up_sub_fram, text="包含子目录", variable=self.include_sub_dir,
                                           command=lambda: self.check_sub_dir())
        self.checkbutton.grid(row=0, column=3, padx=5, pady=5, sticky="w")

        self.label_file_nmb = tk.Label(up_sub_fram, text='文件数')
        self.label_file_nmb.config(font=('Microsoft YaHei UI', 10))
        self.label_file_nmb.grid(row=0, column=4, padx=5, pady=5, sticky="w")

        # list_files_frame = ttk.LabelFrame(self, text="Word旧格式.doc文件, 可多选")
        list_files_frame = ttk.Frame(self)
        list_files_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=10, ipadx=10, sticky='nsew')

        cols = ("文件名", "目录")
        self.treeview = ttk.Treeview(list_files_frame, show="headings", columns=cols, height=13)
        self.treeview.column("# 1", anchor="w", width=428)
        self.treeview.heading("# 1", text="文件名", anchor="w")
        self.treeview.column("# 2", anchor="w", width=288)
        self.treeview.heading("# 2", text="目录", anchor="w")
        # self.treeview.bind("<Return>", lambda e: self.select_multiple())
        # self.treeview.bind('<ButtonRelease-1>', self.cur_select_item)
        # self.treeview.bind('<ButtonRelease-1>', self.select_multiple)
        self.treeview.bind('<Delete>', self.delete_items)
        self.treeview.pack(expand=True, fill='both')

        tree_y_scroll = ttk.Scrollbar(self.treeview, orient='vertical', command=self.treeview.yview)
        self.treeview.configure(yscrollcommand=tree_y_scroll.set)
        tree_y_scroll.place(relx=1, rely=0, relheight=1, anchor='ne')
        # mousewheel scrolling
        self.treeview.bind('<MouseWheel>', lambda event: self.treeview.yview_scroll(-int(event.delta / 60), "units"))

        tree_x_scroll = ttk.Scrollbar(self.treeview, orient='horizontal', command=self.treeview.xview)
        self.treeview.configure(xscrollcommand=tree_x_scroll.set)
        tree_x_scroll.place(relx=0, rely=1, relwidth=1, anchor='sw')
        # event to scroll left / right on Ctrl + mousewheel
        self.treeview.bind('<Control MouseWheel>',
                           lambda event: self.treeview.xview_scroll(-int(event.delta / 60), "units"))

        list_down_frame = ttk.Frame(self)
        list_down_frame.grid(row=3, column=0, columnspan=3, padx=20, pady=10)

        start_remove_button = tk.Button(list_down_frame, text="开始转换", command=self.start_convert_process,
                                        bg='purple',
                                        fg='white',
                                        width=20,
                                        font=('Microsoft YaHei UI', 11, 'bold'))
        start_remove_button.grid(row=0, column=0, columnspan=3, padx=10, pady=5, ipadx=10, ipady=5)

        self.read_all_item_to_treeview_list()

        geekerWin.protocol("WM_DELETE_WINDOW", self.on_window_close)

    def author_frame(self):
        author_frame = ttk.LabelFrame(self, text="关于")
        # author_frame = ttk.Frame(self)
        author_frame.grid(row=4, column=0, columnspan=3, padx=10, pady=10, ipadx=10, sticky='nsew')

        # separator = ttk.Separator(author_frame)
        # separator.grid(row=0, column=0, padx=(20, 10), pady=10, sticky="ew")

        self.mode_switch = ttk.Checkbutton(
            author_frame, text="暗黑/明亮", style="Switch", command=self.toggle_mode, width=60)
        self.mode_switch.grid(row=0, column=0, padx=5, pady=10, sticky="nsew")

        # author_sub_frame = ttk.LabelFrame(author_frame, text="关于sub")
        # author_sub_frame.grid(row=0, column=1, padx=10, pady=10, ipadx=10, sticky='e')

        label_ver = ttk.Label(author_frame, text='Ver 1.0.2.0', font=('Microsoft YaHei UI', 10), cursor="heart")
        label_ver.config(font=('Microsoft YaHei UI', 10))
        label_ver.bind("<Button-1>",
                       lambda e: self.open_website("https://gitee.com/itgeeker/itgeeker_convert_doc_to_docx"))
        label_ver.grid(row=0, column=1, padx=(50, 10), ipadx=10, ipady=10, sticky="nsew")

        label_link = ttk.Label(author_frame, text='www.ITGeeker.net', font=('Microsoft YaHei UI', 10), cursor="hand2")
        label_link.bind("<Button-1>", lambda e: self.open_website("https://www.itgeeker.net"))
        label_link.grid(row=0, column=2, padx=(20, 0), ipadx=10, ipady=10, sticky="nsew")


if __name__ == "__main__":
    icon_b64 = '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'
    icondata = base64.b64decode(icon_b64)
    tmp_p = tempfile.gettempdir()
    tempFile = os.path.join(tmp_p, "icon.ico")
    iconfile = open(tempFile, "wb")
    iconfile.write(icondata)
    iconfile.close()

    geekerWin = tk.Tk()
    geekerWin.wm_iconbitmap(tempFile)
    ## Delete the tempfile
    # os.remove(tempFile)

    style = ttk.Style(geekerWin)
    geekerWin.tk.call("source", "forest-light.tcl")
    geekerWin.tk.call("source", "forest-dark.tcl")
    style.theme_use("forest-dark")

    # start window in the middle of the screen
    # geekerWin.geometry("820x680")
    # geekerWin.eval('tk::PlaceWindow . center')
    window_width = 820
    window_height = 680
    display_width = geekerWin.winfo_screenwidth()
    display_height = geekerWin.winfo_screenheight()

    left = int(display_width / 2 - window_width / 2)
    top = int(display_height / 2 - window_height / 2)
    geekerWin.geometry(f'{window_width}x{window_height}+{left}+{top}')

    # geekerWin.iconbitmap(default=r'geekercloud_orange32.ico')
    # geekerWin.iconbitmap(default=r'D:\git_geeker\geeker_pyinstaller\geekercloud_orange32.ico')

    geekerWin.title("技术奇客小工具-Office Word旧格式(.doc)转换为新格式(.docx)")

    geekerWin.rowconfigure(0, weight=1)
    geekerWin.columnconfigure(0, weight=1)

    app_convert_doc = AppConvertDoc(geekerWin)
    app_convert_doc.pack()
    geekerWin.mainloop()
