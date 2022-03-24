from tools import crawling_SE, operate_excel
from tkinter import *
from tkinter.filedialog import askopenfilename
import threading

SEARCH = crawling_SE
EXCEL = operate_excel

# func

def define_layout(obj, cols=1, rows=1):

    def method(trg, col, row):

        for c in range(cols):
            trg.columnconfigure(c, weight=1)
        for r in range(rows):
            trg.rowconfigure(r, weight=1)

    if type(obj) == list:
        [method(trg, cols, rows) for trg in obj]
    else:
        trg = obj
        method(trg, cols, rows)


def selectPath():
    path_ = askopenfilename(title="选择 xlsx 表格文件", defaultextension="xlsx")
    path.set(path_)


def get_excel_kws_research():
    msg.set("正在运行...")
    btn_run_section_1["state"] = "disabled"
    window.update()
    filepath = input_filepath_section_1.get()
    if(filepath == ''):
        msg.set("请选择/输入路径!")
        btn_run_section_1["state"] = "normal"
        return
    wbname = input_wbname_section_1.get()
    if(wbname == ''):
        wbname = 'Sheet1'
    td = input_td_section_1.get()
    if(td == ''):
        td = 'A2'
    app = EXCEL.init_data_hide()
    try:
        msg.set("正在读取表格...")
        window.update()
        col = td[0]
        row = td[1]
        wb1 = EXCEL.get_wb(app, filepath)
        sht1 = EXCEL.get_sht(wb1, wbname)
        kws_list = EXCEL.get_td_value(sht1, col, row)
        app.quit()
        if(isinstance(kws_list, list) == False):
            kws_list = [kws_list]
    except:
        msg.set("运行失败:表格读取失败")
        app.quit()
        btn_run_section_1["state"] = "normal"
        return
    if(kws_list[0] == None):
        msg.set("运行失败:无数据")
        btn_run_section_1["state"] = "normal"
        return
    try:
        msg.set("正在爬取..请等待...")
        window.update()
        val = SEARCH.get_kws_research_more(
            kws_list, SEARCH.search_google_result)
    except:
        msg.set("运行失败:爬取结果失败.")
        btn_run_section_1["state"] = "normal"
        return
    app = EXCEL.init_data_hide()
    try:
        msg.set("正在写入表格..请等待...")
        wb1 = EXCEL.get_wb(app, filepath)
        sht2 = EXCEL.add_sht(wb1)
        window.update()
        if(val):
            EXCEL.init_th(sht2, ['keyword', 'title',
                                 'link', 'domain', 'description'])
            for it in val:
                EXCEL.write_td(
                    sht2, it, f"a{int(EXCEL.get_count_all_rows(sht2))+1}")
        EXCEL.get_rng(sht2, 'b2:e2', EXCEL.style_autofit_col)
        wb1.save()
        msg.set("运行成功!已生成新工作表格,可打开文件查看.")
    except:
        msg.set("运行失败:写入表格失败.")
    app.quit()
    btn_run_section_1["state"] = "normal"
    window.update()
    return


def threadIt(func, *args):
    # 将函数打包进线程
    t = threading.Thread(target=func, args=args)
    t.setDaemon(True)
    t.start()


# style set
color_bg = '#FFFFFF'
color_main_font = '#424242'
color_sub_font = '#A4A4A4'
color_main = '#01A9DB'
color_sub = '#F5BCA9'
color_input = '#E6E6E6'
color_btn = '#424242'
color_btn_font = '#FFFFFF'
font_style_1 = 'microsoft yahei'

# set
align_mode = 'nswe'
pad = 5
div_size = 200
max_col = 3

# 窗口
if __name__ == '__main__':
    window = Tk()
    window.focus_force()
    window_x, window_y = window.maxsize()
    window_width = 500
    window_height = 300
    window_x = (window_x-window_width)/2
    window_y = (window_y-window_height)/2
    window.title('Google Crawler')
    window.geometry("%dx%d+%d+%d" %
                    (window_width, window_height, window_x, window_y))
    window['bg'] = color_bg
    # window.attributes('-topmost', 1)

    # frame
    section_1 = Frame(window,  width=div_size, height=300, bg=color_bg)
    section_1_div1 = Frame(section_1, bg=color_bg)
    section_1_div2 = Frame(section_1, bg=color_bg)
    section_1_div3 = Frame(section_1, bg=color_bg)

    section_2 = Frame(window,  width=div_size, height=50, bg=color_bg)
    section_2_div1 = Frame(section_2, bg=color_bg)

    section_3 = Frame(window,  width=div_size, height=50, bg='orange')
    section_4 = Frame(window,  width=div_size, height=50, bg='green')

    section_1.grid(row=0, column=0, columnspan=max_col,
                   padx=pad, pady=pad, sticky=align_mode)
    section_1_div1.grid(padx=pad, pady=pad, row=0)
    section_1_div2.grid(padx=pad, pady=pad, row=1)
    section_1_div3.grid(padx=pad, pady=pad, row=2)

    section_2.grid(row=1, column=0, columnspan=max_col,
                   padx=pad, pady=pad, sticky=align_mode)
    section_2_div1.grid(padx=pad, pady=pad, row=0, sticky=W)

    # 元素
    path = StringVar()
    msg = StringVar()
    msg.set("为避免Google封IP,增加了间隔时间,请耐心等待")
    # section 1
    Label(section_1_div1, text="Google Crawling", font=(font_style_1, 14, 'bold'),
          bg=color_bg, fg=color_main_font).grid(column=0, columnspan=max_col, row=0)

    Label(section_1_div1, font=(font_style_1, 8), text="使用说明: 程序会查找xlsx表格中的工作表所选单元格列的所有数据,根据数据爬取搜索结果，在表格中创建新工作表导出数据.(默认根据网络所在的地区搜索结果)", wraplength=400, justify='left', bg=color_bg, fg=color_main).grid(
        column=0, columnspan=max_col, row=1)

    Label(section_1_div2, text="文件路径:", font=(font_style_1, 10), bg=color_bg, fg=color_main_font).grid(
        column=0, row=2, sticky=E)

    input_filepath_section_1 = Entry(
        section_1_div2, textvariable=path, font=(font_style_1, 10), bg=color_input, fg=color_main_font)
    input_filepath_section_1.grid(column=1, row=2)

    Button(
        section_1_div2, text=">> 选择文件 ", command=selectPath, font=(font_style_1, 8), bg=color_btn, fg=color_btn_font).grid(column=2, row=2, sticky=W)

    Label(section_1_div2, text="工作表> 表名:", font=(font_style_1, 10), bg=color_bg, fg=color_main_font).grid(
        column=0, row=3, sticky=E)

    input_wbname_section_1 = Entry(
        section_1_div2, bg=color_input, fg=color_main_font, font=(font_style_1, 10))
    input_wbname_section_1.grid(column=1, row=3)

    Label(section_1_div2, font=(font_style_1, 8), text="不填默认:Sheet1", bg=color_bg, fg=color_sub_font).grid(
        column=2, row=3, sticky=W)

    Label(section_1_div2, text="单元格> 列&行:", font=(font_style_1, 10), bg=color_bg, fg=color_main_font).grid(
        column=0, row=4, sticky=E)

    input_td_section_1 = Entry(
        section_1_div2, bg=color_input, fg=color_main_font, font=(font_style_1, 10))
    input_td_section_1.grid(column=1, row=4)

    Label(section_1_div2, font=(font_style_1, 8), text="不填默认:A2", bg=color_bg, fg=color_sub_font).grid(
        column=2, row=4, sticky=W)

    Label(section_1_div3, font=(font_style_1, 9), textvariable=msg, wraplength=350, justify='left', bg=color_bg, fg=color_main_font).grid(
        column=0, columnspan=max_col, row=0)

    btn_run_section_1 = Button(
        section_1_div3, text=" -> 开始运行 <- ", command=lambda: threadIt(get_excel_kws_research), font=(font_style_1, 10), bg='#0489B1', fg=color_btn_font)
    btn_run_section_1.grid(column=0, columnspan=max_col, row=1)

    Label(section_2_div1, text="更新: 增加结果数至20条;添加GUI;添加表格配置;编译EXE程序;压缩程序大小;多线程运行避免卡死;", font=(font_style_1, 8), wraplength=400, justify='left', bg=color_bg, fg=color_sub_font).grid(
        column=0, columnspan=max_col, row=0)
    # else
    define_layout(window, cols=2, rows=3)
    define_layout([section_1, section_2, section_3, section_4])

    # 结束
    window.update()
    window.mainloop()
    print('>> done <<')

# 爬取速度：10个关键词共200条数据，用时5分46秒
