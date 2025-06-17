"""
   ______  ____  __   ____  _             _ _
  / ___\ \/ /  \/  | / ___|| |_ _   _  __| (_) ___
 | |    \  /| |\/| | \___ \| __| | | |/ _` | |/ _ \
 | |___ /  \| |  | |  ___) | |_| |_| | (_| | | (_) |
  \____/_/\_\_|  |_| |____/ \__|\__,_|\__,_|_|\___/

开发作者：晨小明
开发日期：2024/01/04
开发版本：v13.0__release
发布版本：v4.0__release
修改日期：2025/06/17
主要功能：一、支持单文件处理或批量文档处理，输入文件路径或文件夹路径，自动判断。
         二、读取.docx文件并设置格式：
        三、支持添加页码（可选）：4号半角宋体阿拉伯数字，数字左右各加一条4号“一字线”，奇数页在右侧左空一字，偶数页在左侧左空一字
        四、识别文档中的图片并输出（可选）：（注：图片可能会被压缩）
        五、替换功能
            1.符号替换
                将英文状态下的符号替换为中文状态下的相同符号，包含如下：
                "(" --> "（"
                ")" --> "）"
                ")、" --> "）"
                "）、" --> "）"
                "," --> "，"
                ":" --> "："
                ";" --> "；"
                "?" --> "？"
                " " --> ""
            2.其他格式
                数字后有顿号替换为点，如："1、" --> "1."
         六、输出文件名称含时间点，方便标记（可选）
         （注，本程序无法处理图片格式，如果图片独立成段，本程序所用API识别到图片会被默认是空段落，为了防止图片删除，只能放弃处理空段落及图片格式）
更新日志：
【新增】支持用户手动输入路径，输入类型多样化；
【新增】底部版本信息；
【新增】全角空格替换；
【新增】左侧缩进为0（不是首行缩进）；
【新增】段前段后为0；
【新增】取消孤行控制；
【优化】界面排版优化，视觉效果更佳；
【修复】两位数字后为顿号（、）时，会丢失相邻数之前的数字；
【修复】其他问题。
"""

from docx import Document
from docx.shared import Pt, Cm  # 用来设置字体的大小
from docx.oxml.ns import qn  # 设置字体
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT  # 设置对其方式
from docx.oxml import OxmlElement
from os import listdir, path, makedirs, getcwd
from tkinter import Tk, Entry, Button, Label, filedialog, messagebox, SUNKEN, Radiobutton, Frame, ttk, Listbox, StringVar, END, Toplevel, Canvas
from time import localtime, strftime
from PIL import Image, ImageTk
from webbrowser import open as webopen


def margin(docx):
    """ 设置页边距 """
    for s in docx.sections:
        s.top_margin = Cm(PAGETOPMARGIN)
        s.bottom_margin = Cm(PAGEBOTTOMMARGIN)
        s.left_margin = Cm(PAGELEFTMARGIN)
        s.right_margin = Cm(PAGERIGHTMARGIN)


def footer(docx):
    """ 设置页脚，添加页码 """
    # print(len(docx.sections))
    def AddFooterNumber(p):
        t1 = p.add_run("— ")
        font = t1.font
        font.name = PAGENUMBERFONT
        font.size = Pt(PAGENUMBERFONTSIZE)  # 14号字体
        t1._element.rPr.rFonts.set(qn("w:eastAsia"), PAGENUMBERFONT)

        run1 = p.add_run('')
        fldChar1 = OxmlElement('w:fldChar')  # creates a new element
        fldChar1.set(qn('w:fldCharType'), 'begin')  # sets attribute on element
        run1._element.append(fldChar1)

        run2 = p.add_run('')
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')  # sets attribute on element
        instrText.text = 'PAGE'
        font = run2.font
        font.name = PAGENUMBERFONT
        font.size = Pt(PAGENUMBERFONTSIZE)  # 14号字体
        run2._element.rPr.rFonts.set(qn("w:eastAsia"), PAGENUMBERFONT)
        run2._element.append(instrText)

        run3 = p.add_run('')
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run3._element.append(fldChar2)

        t2 = p.add_run(" —")
        font = t2.font
        font.name = PAGENUMBERFONT
        font.size = Pt(PAGENUMBERFONTSIZE)  # 14号字体
        t2._element.rPr.rFonts.set(qn("w:eastAsia"), PAGENUMBERFONT)

    for s in docx.sections:
        # print(s.footer)
        footer = s.footer  # 获取第一个节的页脚
        footer.is_linked_to_previous = True  # 编号续前一节
        paragraph = footer.paragraphs[0]  # 获取页脚的第一个段落
        paragraphFun("odd_footer", paragraph)
        AddFooterNumber(paragraph)
        even_footer = s.even_page_footer  # 获取第一个节的页脚
        even_footer.is_linked_to_previous = True  # 编号续前一节
        paragraph = even_footer.paragraphs[0]  # 获取页脚的第一个段落
        paragraphFun("even_footer", paragraph)
        AddFooterNumber(paragraph)


def paragraphFun(is_title, p):
    """ 段落函数 """
    try:
        p.paragraph_format.element.pPr.ind.set(qn("w:leftChars"), '0')  # 左侧缩进
        p.paragraph_format.element.pPr.ind.set(qn("w:rightChars"), '0')  # 右侧缩进
        p.paragraph_format.element.pPr.ind.set(qn("w:left"), '0')  # 缩进(cm)
        p.paragraph_format.element.pPr.ind.set(qn("w:right"), '0')  # 缩进(cm)
    except:
        pass
    if is_title == "title":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = Pt(TITLEMARGIN)  # 行距
        p.paragraph_format.first_line_indent = None
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
    elif is_title == "odd_footer":
        p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        p.paragraph_format.right_indent = Pt(14)
        p.paragraph_format.line_spacing = Pt(TEXTMARGIN)
    elif is_title == "even_footer":
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        p.paragraph_format.left_indent = Pt(14)
        p.paragraph_format.line_spacing = Pt(TEXTMARGIN)
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = Pt(TEXTMARGIN)  # 行距
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.widow_control = False


def isLevel1(p):
    """ 判断是否是 1 级标题 """
    index1_list = ["一、", "二、", "三、", "四、", "五、", "六、", "七、", "八、", "九、", "十、",
                   "十一、", "十二、", "十三、", "十四、", "十五、", "十六、", "十七、", "十八、", "十九、", "二十、"]
    for i in range(len(index1_list)):
        if p.text.find(index1_list[i]) != -1:
            if '。' in p.text:
                p.text = p.text.replace('。', '')
            if '？' in p.text:
                p.text = p.text.replace('？', '')
            if '：' in p.text:
                p.text = p.text.replace('：', '')
            if '；' in p.text:
                p.text = p.text.replace('；', '')
            return "level1"
        else:
            continue


def isLevel2(p):
    """ 判断是否是 2 级标题 """
    index2_list = ["（一）", "（二）", "（三）", "（四）", "（五）", "（六）", "（七）", "（八）", "（九）", "（十）",
                   "（十一）", "（十二）", "（十三）", "（十四）", "（十五）", "（十六）", "（十七）", "（十八）", "（十九）", "（二十）"]
    for i in range(len(index2_list)):
        if p.text.find(index2_list[i]) != -1:
            if '。' in p.text:
                p.text = p.text.replace('。', '')
            if '？' in p.text:
                p.text = p.text.replace('？', '')
            if '：' in p.text:
                p.text = p.text.replace('：', '')
            if '；' in p.text:
                p.text = p.text.replace('；', '')
            return "level2"
        else:
            continue


def text(is_title, is_level1, is_level2, is_digit, p, i, version_ipt):
    """ 正文函数 """
    if is_title == "title":
        if is_digit == "num_or_let":
            title_num = p.add_run(i)
            title_num.font.name = NUMBERFONT
            if version_ipt == "school":
                title_num.font.size = Pt(FONTSIZEDICT["小二号"])
            else:
                title_num.font.size = Pt(FONTSIZEDICT["二号"])
            title_num.font.bold = False
        else:
            run_title = p.add_run(i)
            run_title.font.name = TITLEFONT
            run_title._element.rPr.rFonts.set(qn('w:eastAsia'), TITLEFONT)
            if version_ipt == "school":
                run_title.font.size = Pt(FONTSIZEDICT["小二号"])
            else:
                run_title.font.size = Pt(FONTSIZEDICT["二号"])
            run_title.font.bold = False
    else:
        if is_digit == "num_or_let":
            text_num = p.add_run(i)
            text_num.font.name = NUMBERFONT
            if version_ipt == "school":
                text_num.font.size = Pt(FONTSIZEDICT["四号"])
            else:
                text_num.font.size = Pt(FONTSIZEDICT["三号"])
            text_num.font.bold = False
        elif is_level1 == "level1":
            run_level1 = p.add_run(i)
            run_level1.font.name = LEVEL1FONT
            run_level1._element.rPr.rFonts.set(qn('w:eastAsia'), LEVEL1FONT)
            if version_ipt == "school":
                run_level1.font.size = Pt(FONTSIZEDICT["四号"])
            else:
                run_level1.font.size = Pt(FONTSIZEDICT["三号"])
            run_level1.font.bold = False
        elif is_level2 == "level2":
            run_level2 = p.add_run(i)
            run_level2.font.name = LEVEL2FONT
            run_level2._element.rPr.rFonts.set(qn('w:eastAsia'), LEVEL2FONT)
            if version_ipt == "school":
                run_level2.font.size = Pt(FONTSIZEDICT["四号"])
            else:
                run_level2.font.size = Pt(FONTSIZEDICT["三号"])
            run_level2.font.bold = False
        else:
            run_content = p.add_run(i)
            run_content.font.name = TEXTFONT
            run_content._element.rPr.rFonts.set(qn('w:eastAsia'), TEXTFONT)
            if version_ipt == "school":
                run_content.font.size = Pt(FONTSIZEDICT["四号"])
            else:
                run_content.font.size = Pt(FONTSIZEDICT["三号"])
            run_content.font.bold = False


def replace(p):
    """ 替换函数 """
    # 替换符号
    if ')、' in p.text:
        p.text = p.text.replace(')、', '）')
    if '）、' in p.text:
        p.text = p.text.replace('）、', '）')
    if '(' in p.text:
        p.text = p.text.replace('(', '（')
    if ')' in p.text:
        p.text = p.text.replace(')', '）')
    if ',' in p.text:
        p.text = p.text.replace(',', '，')
    if ':' in p.text:
        p.text = p.text.replace(':', '：')
    if ';' in p.text:
        p.text = p.text.replace(';', '；')
    if '?' in p.text:
        p.text = p.text.replace('?', '？')
    if ' ' in p.text:  # 空格
        p.text = p.text.replace(' ', '')
    if '　' in p.text:  # U+3000
        p.text = p.text.replace('　', '')
    if ' ' in p.text:  # U+2003
        p.text = p.text.replace(' ', '')
    return p


def fixDocx(docx, version_ipt):
    """ 主要格式 """
    lvl = 0
    for p in docx.paragraphs:
        # print(p.text)
        if p.text == '':
            continue
        else:
            lvl += 1
            # print(p.paragraph_format.first_line_indent)
            # print(p.style.font.size)
            p = replace(p)
            is_level1 = isLevel1(p)
            is_level2 = isLevel2(p)
            if lvl == 1:
                paragraphFun("title", p)
                for run_title in p.runs:
                    # print(run_title.text)
                    string = run_title.text
                    # print(string)
                    run_title._element.getparent().remove(run_title._element)
                    for i in string:
                        num_or_let = isNumberOrLetter(i)
                        text("title", is_level1, is_level2, num_or_let, p, i, version_ipt)
            else:
                paragraphFun("text", p)
                for run_content in p.runs:
                    # print(run_content.text)
                    string = run_content.text
                    # print(string)
                    p.paragraph_format.first_line_indent = 0  # 首行缩进
                    p.paragraph_format.element.pPr.ind.set(qn("w:firstLineChars"), '200')  # 首行缩进
                    run_content._element.getparent().remove(run_content._element)
                    # 替换格式："1、" --> "1."
                    s = ''
                    for i in range(len(string)):
                        if i + 1 <= len(string):
                            if string[i:i+1] in '0123456789' and string[i+1:i+2] == "、":
                                # print(string[i:i+1])
                                str = string[i:i+1] + "." + string[i+2:]
                                s += str
                                break
                            else:
                                s += string[i:i+1]
                    for i in s:  # 遍历字符串
                        num_or_let = isNumberOrLetter(i)
                        text("notitle", is_level1, is_level2, num_or_let, p, i, version_ipt)


def isNumberOrLetter(char):
    """ 判断是否为数字或字母 """
    number_and_letter_strs = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'
    if char in number_and_letter_strs:
        return "num_or_let"
    else:
        return False


def picFix(docx, file, output_path):
    """ 图片处理 """
    img_path = output_path + "\image"
    file_name = path.splitext(file)[0]
    parts = docx.part.related_parts
    parts_values = parts.values()
    parts_keys = parts.keys()
    list_val = list(parts_values)
    list_key = list(parts_keys)
    parts_length = len(parts_values)
    if parts_length > 5:
        # print(type(list(parts_values)[-1]))
        k = 0
        for i in range(parts_length):
            # print(type(list_val[i]))
            if 'image' in str(type(list_val[i])):
                if not path.isdir(img_path):
                    makedirs(img_path)
                # print('找到图片数据')
                k += 1
                try:
                    img_data = parts[list_key[i]].image.blob
                    img_type = parts[list_key[i]].image.ext
                    full_path = f'{img_path}\{file_name}_image{k}.{img_type}'
                    print(f"··>提示<·· 正在输出：{full_path}")
                    with open(full_path, 'wb') as f:
                        f.write(img_data)
                except:
                    print(f"··>错误<·· 图片{k}输出失败！")
        if k == 0:
            print(f"··>提示<·· 未找到图片！")


def fixWord(docx_path, save_path, file, output_path, time_ipt, page_ipt, img_ipt, version_ipt):
    """ 文档处理 """
    docx = Document(docx_path)

    # 页边距
    margin(docx)

    # 修改格式
    fixDocx(docx, version_ipt)

    # 添加时间后缀
    file_name = path.splitext(file)[0]
    if time_ipt == "1":
        save_time = strftime("%m%d%H%M", localtime())
        save_path = output_path + f"\{file_name}" + save_time + ".docx"
    else:
        save_time = ""
        save_path = output_path + f"\{file_name}" + ".docx"

    # 设置页码
    if page_ipt == "1":
        # 奇偶页不同
        docx.settings.odd_and_even_pages_header_footer = True
        footer(docx)

    # 保存文档中的图片
    if img_ipt == "1":
        picFix(docx, file, output_path)
    output_time = strftime("%m-%d %H:%M:%S", localtime())
    docx.save(save_path)
    output_txt = output_time + "    " + save_path
    play_history_frm_listbox.insert(END, output_txt)
    print(f"··>提示<·· 已保存：{output_txt}")
    # 设置滚动条位置到最大值，即拖动到最底部
    play_history_frm_listbox.yview_moveto(1)
    # play_history_frm_listbox.xview_moveto(1)
    return save_time


def inputPath():
    """ 输入路径 """
    input_path = type_radio_value.get()
    if input_path == "file_path":
        file_path = filedialog.askopenfile(title="请选择文件", filetypes=[("docx文件", "*.docx")])
        if file_path != None:
            path_entry.delete(0, END)
            path_entry.insert(0, file_path.name)
    elif input_path == "dir_path":
        dir_path = filedialog.askdirectory(title="请选择文件夹")
        if dir_path != "":
            path_entry.delete(0, END)
            path_entry.insert(0, dir_path)


def versionText1():
    """ 版本文本控制1 """
    version_text_label["text"] = VERSIONTEXT1


def versionText2():
    """ 版本文本控制2 """
    version_text_label["text"] = VERSIONTEXT2


def inputFile():
    """ 选择文件 """
    browse_path_button.config(text="选择文件")
    path_entry.delete(0, END)


def inputDir():
    """ 选择文件夹 """
    browse_path_button.config(text="选择文件夹")
    path_entry.delete(0, END)


def reSet():
    """ 重置 """
    type_radio1.select()
    inputFile()
    version_radio1.select()
    versionText1()
    time_radio2.select()
    page_radio2.select()
    img_radio2.select()
    play_history_frm_listbox.delete(0, END)


def on_enter(event):
    event.widget.config(cursor="hand2", fg="blue")


def on_leave(event):
    event.widget.config(cursor="", fg="#888888")


def toMail(event):
    """ 打开邮箱 """
    webopen("mailto:3038693133@qq.com")


def wxTk(event):
    wx_tk = Toplevel(tk)
    original_image = Image.open(wxgzh_path)
    wx_tk.geometry(f"{original_image.width}x460+0+0")
    wx_tk.iconbitmap(icon_path)
    wx_tk.title("微信公众号：晨小明工作室")
    wx_title = Label(wx_tk, text="微信扫一扫关注公众号", font=("微软雅黑", 14))
    wx_title.grid(row=0, column=0, padx=2, pady=2)
    # 创建Canvas
    cv = Canvas(wx_tk, width=original_image.width, height=original_image.height + 30, highlightthickness=0)
    cv.grid(row=1, column=0, padx=2, pady=0)
    # 加载图片
    time_icon = original_image.resize((round(original_image.width / 1), round(original_image.height / 1)))  # 缩放图片到指定大小
    time_icon_new = ImageTk.PhotoImage(time_icon)
    cv.create_image(0, 0, image=time_icon_new, anchor="nw")
    wx_tk.mainloop()


def checkPath():
    """ 检查路径 """
    input_path = path_entry.get()
    if input_path == "":
        messagebox.showerror("错误", "请选择文件或文件夹路径！")
    else:
        file_type = type_radio_value.get()
        if file_type == "file_path":
            if path.isfile(input_path):
                # print("··>提示<·· 检查路径成功！")
                return True
            else:
                messagebox.showerror("错误", "文件路径错误！")
        elif file_type == "dir_path":
            if path.isdir(input_path):
                # print("··>提示<·· 检查路径成功！")
                return True
            else:
                messagebox.showerror("错误", "文件夹路径错误！")
    return False


def main():
    """ 主函数 """
    if checkPath():
        input_path = path_entry.get().replace("/", "\\")
        version_ipt = version_radio_value.get()
        file_type = type_radio_value.get()
        time_ipt = time_radio_value.get()
        page_ipt = page_radio_value.get()
        img_ipt = img_radio_value.get()
        if file_type == "dir_path":
            output_path = input_path + "\output"
            have_docx = 0
            for file in listdir(input_path):
                if '~' in file:
                    continue
                elif file.endswith('.docx'):
                    if not path.isdir(output_path):
                        makedirs(output_path)
                    have_docx += 1
                    file_path = path.join(input_path, file)
                    fixWord(file_path, output_path + f'\{file}', file, output_path, time_ipt, page_ipt, img_ipt, version_ipt)
            if have_docx == 0:
                print("··>错误<·· 没有找到.docx文件")
                messagebox.showwarning("警告", "没有找到.docx文件！")
            else:
                messagebox.showinfo("提示", "全部处理完成！\n输出路径：" + output_path)
        elif file_type == "file_path":
            # 文件名
            file = input_path.split("\\")[-1]
            # 输出路径
            dir_path = input_path.split("\\")
            dir_path.pop()
            result = '\\'.join(str(x) for x in dir_path)
            output_path = result + "\output"
            if not path.isdir(output_path):
                makedirs(output_path)
            save_time = fixWord(input_path, output_path + f'\{file}', file, output_path, time_ipt, page_ipt, img_ipt, version_ipt)
            messagebox.showinfo("提示", "处理完成！\n输出路径：" + output_path + "\\" + file.split(".")[0] + save_time + ".docx")


if __name__ == '__main__':
    VERSION = "v4.0"
    UPDATETIME = "2025年6月17日"
    """
        !!!!!!!!!!!!
        打包时把此路径改为相对路径，并把图图片复制粘贴到打包后的根目录里
        !!!!!!!!!!!!
    """
    icon_path = getcwd() + "\\static\\icon.ico"
    wxgzh_path = getcwd() + "\\static\\wxgzh.jpg"
    cxm_path = getcwd() + "\\static\\cxmstudio-lignt-heng.png"
    # 配置信息start
    # 字号字典
    FONTSIZEDICT = {
        "八号": 5, "七号": 5.5, "小六号": 6.5, "六号": 7.5, "小五号": 9, "五号": 10.5, "小四号": 12, "四号": 14, "小三号": 15, "三号": 16, "小二号": 18, "二号": 22, "小一号": 24, "一号": 26, "小初号": 36, "初号": 42
    }
    # 版本文本
    VERSIONTEXT1 = """当前配置：
    版     本：学校留存
    页 边  距：上3.7cm 下3.5cm 左2.8cm 右2.6cm
    标     题：小二号 方正小标宋简体  33磅
    正     文：四号 仿宋_GB2312 28磅
    一级 标题：四号 黑体 28磅
    二级 标题：四号 楷体_GB2312 28磅
    数字&英文：四号 TimesNewRoman字体
    页     码：四号 宋体
"""
    VERSIONTEXT2 = """当前配置：
    版     本：上交上报
    页 边  距：上3.7cm 下3.5cm 左2.8cm 右2.6cm
    标     题：二号 方正小标宋简体  33磅
    正     文：三号 仿宋_GB2312 28磅
    一级 标题：三号 黑体 28磅
    二级 标题：三号 楷体_GB2312 28磅
    数字&英文：三号 TimesNewRoman字体
    页     码：四号 宋体
"""
    # 页边距-CM
    PAGETOPMARGIN = 3.7  # 页边距：上3.7cm
    PAGEBOTTOMMARGIN = 3.5  # 页边距：下3.5cm
    PAGELEFTMARGIN = 2.8  # 页边距：左2.8cm
    PAGERIGHTMARGIN = 2.6  # 页边距：右2.6cm
    # 标题-磅值
    TITLEFONT = '方正小标宋简体'  # 标题：方正小标宋简体
    TITLEMARGIN = 33  # 标题：固定值33磅
    # 正文-磅值
    TEXTFONT = '仿宋_GB2312'  # 正文：仿宋_GB2312
    TEXTMARGIN = 28  # 正文：一般固定值28磅
    LEVEL1FONT = '黑体'  # 一级标题：黑体
    LEVEL2FONT = '楷体_GB2312'  # 二级标题：楷体_GB2312
    NUMBERFONT = 'Times New Roman'  # 数字&英文：TimesNewRoman字体
    # 页码-磅值
    PAGENUMBERFONTSIZE = FONTSIZEDICT["四号"]  # 页码：四号
    PAGENUMBERFONT = '宋体'  # 页码：宋体
    # 配置信息end
    # tkinter start
    tk = Tk()
    tk.title(f"文档处理工具 {VERSION} （学校定制版）")
    screen_width = tk.winfo_screenwidth()
    screen_height = tk.winfo_screenheight()
    tk.iconbitmap(icon_path)
    tk.geometry("1000x580")
    # 刷新窗口参数
    tk.update()
    # 计算窗口居中时左上角的坐标
    x = (screen_width - tk.winfo_width()) // 2
    y = (screen_height - tk.winfo_height()) // 2
    tk.geometry(f"+{x}+{y-30}")
    # tk.attributes("-alpha", 0.8)
    windw_width = tk.winfo_width()
    windw_height = tk.winfo_height()
    # 输入类型单选按钮
    type_frm = Frame(tk)
    type_frm.pack()
    # 空占位符
    label_ = Label(type_frm, font=("Ya Hei", 10), text=" ")
    label_.grid(row=0, column=0, padx=5, pady=5, sticky="e")
    type_label = Label(type_frm, font=("Ya Hei", 10, "bold"), text="请选择输入类型：")
    type_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
    type_radio_value = StringVar()
    type_radio1 = Radiobutton(type_frm, text="文件", font=("Ya Hei", 10), value="file_path", variable=type_radio_value, command=inputFile)
    type_radio1.grid(row=1, column=1, padx=5, pady=2)
    type_radio2 = Radiobutton(type_frm, text="文件夹", font=("Ya Hei", 10), value="dir_path", variable=type_radio_value, command=inputDir)
    type_radio2.grid(row=1, column=2, padx=5, pady=2)
    type_radio1.select()
    browse_path_button = Button(type_frm, font=("Ya Hei", 10), text="选择文件", command=inputPath)
    browse_path_button.grid(row=1, column=3, padx=5, pady=2)
    # 文件路径
    path_frm = Frame(tk)
    path_frm.pack()
    # 输入框
    path_entry = Entry(path_frm, width=60, font=("Ya Hei", 10), textvariable="222", border=1, relief="solid")
    path_entry.grid(row=0, column=0, padx=5, pady=5, ipadx=2, ipady=4, sticky="w")
    # 分隔线
    separator = Frame(tk, height=2, bd=1, relief=SUNKEN)
    separator.pack(fill="x", padx=5, pady=5)
    # 选择版本
    version_frm = Frame(tk)
    version_frm.pack(side="top", padx=2, pady=2)
    version_choose_frm = Frame(version_frm)
    version_choose_frm.grid(row=0, column=0, padx=5, pady=5)
    version_label = Label(version_choose_frm, font=("Ya Hei", 10, "bold"), text="请选择版本")
    version_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
    version_radio_value = StringVar()
    version_radio1 = Radiobutton(version_choose_frm, font=("Ya Hei", 10), text="学校留存", variable=version_radio_value, value="school", command=versionText1)
    version_radio1.grid(row=1, column=0, padx=5, pady=5)
    version_radio2 = Radiobutton(version_choose_frm, font=("Ya Hei", 10), text="上交上报", variable=version_radio_value, value="report", command=versionText2)
    version_radio2.grid(row=2, column=0, padx=5, pady=5)
    version_radio1.select()
    version_text_frm = Frame(version_frm)
    version_text_frm.grid(row=0, column=1, padx=5, pady=5)
    version_text_label = Label(version_text_frm, width=46, height=12, font=("Ya Hei", 10), text=VERSIONTEXT1, justify="left", border=1, relief="solid")
    version_text_label.grid(row=0, column=1, sticky="w")
    # 处理信息
    infos_frm = Frame(version_frm)
    infos_frm.grid(row=0, column=2, padx=5, pady=5)
    info_frm = Frame(infos_frm)
    info_frm.grid(row=0, column=0, padx=5, pady=5)
    time_label = Label(info_frm, font=("Ya Hei", 10, "bold"), text="添加时间标记：")
    time_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
    time_radio_value = StringVar()
    time_radio1 = Radiobutton(info_frm, font=("Ya Hei", 10), text="是", variable=time_radio_value, value=True)
    time_radio1.grid(row=0, column=1, padx=5, pady=5)
    time_radio2 = Radiobutton(info_frm, font=("Ya Hei", 10), text="否", variable=time_radio_value, value=False)
    time_radio2.grid(row=0, column=2, padx=5, pady=5)
    time_radio2.select()
    page_label = Label(info_frm, font=("Ya Hei", 10, "bold"), text="添加页码：")
    page_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
    page_radio_value = StringVar()
    page_radio1 = Radiobutton(info_frm, font=("Ya Hei", 10), text="是", variable=page_radio_value, value=True)
    page_radio1.grid(row=1, column=1, padx=5, pady=5)
    page_radio2 = Radiobutton(info_frm, font=("Ya Hei", 10), text="否", variable=page_radio_value, value=False)
    page_radio2.grid(row=1, column=2, padx=5, pady=5)
    page_radio2.select()
    img_label = Label(info_frm, font=("Ya Hei", 10, "bold"), text="保存文档中的图片：")
    img_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
    img_radio_value = StringVar()
    img_radio1 = Radiobutton(info_frm, font=("Ya Hei", 10), text="是", variable=img_radio_value, value=True)
    img_radio1.grid(row=2, column=1, padx=5, pady=5)
    img_radio2 = Radiobutton(info_frm, font=("Ya Hei", 10), text="否", variable=img_radio_value, value=False)
    img_radio2.grid(row=2, column=2, padx=5, pady=5)
    img_radio2.select()
    # 分隔线
    separator = Frame(tk, height=2, bd=1, relief=SUNKEN)
    separator.pack(fill="x", padx=5, pady=5)
    # 处理按钮
    btn_frm = Frame(tk)
    btn_frm.pack(pady=6)
    merge_button = Button(btn_frm, font=("Ya Hei", 10, "bold"), text="开始处理", command=main)
    merge_button.grid(row=0, column=0, padx=5, pady=5)
    label_ = Label(btn_frm, font=("Ya Hei", 10), text=" ")
    label_.grid(row=0, column=1, padx=5, pady=5, sticky="e")
    merge_button = Button(btn_frm, font=("Ya Hei", 10), text="重置", fg="blue", command=reSet)
    merge_button.grid(row=0, column=2, padx=5, pady=5)
    # # 虚线分隔线
    # separator_ = Frame(infos_frm)
    # separator_.grid(row=0, column=1, padx=34, pady=5)
    # separator_s = ttk.Separator(infos_frm, orient="vertical")
    # separator_s.grid(row=0, column=1, sticky="ns")
    # 分隔线
    separator = Frame(tk, height=2, bd=1, relief=SUNKEN)
    separator.pack(fill="x", padx=2, pady=2)
    # 处理日志
    play_history_frm = Frame(tk)
    play_history_frm.pack()
    play_history_frm_lbl = Label(play_history_frm, text="处理日志", font=("Ya Hei", 10, "bold"), fg="green")
    play_history_frm_lbl.grid(row=0, column=0, padx=5, pady=5)
    play_history_frm_listbox = Listbox(play_history_frm, width=100, height=6, font=("Ya Hei", 10), border=1, activestyle="none")
    play_history_frm_listbox.grid(row=1, column=0, padx=5, pady=5, ipadx=5, ipady=5)
    play_history_scroll_bar_v = ttk.Scrollbar(play_history_frm, orient="vertical", command=play_history_frm_listbox.yview)
    play_history_scroll_bar_v.grid(row=1, column=1, sticky='ns')
    play_history_scroll_bar_h = ttk.Scrollbar(play_history_frm, orient="horizontal", command=play_history_frm_listbox.xview)
    play_history_scroll_bar_h.grid(row=2, column=0, sticky='we')
    play_history_frm_listbox.configure(yscrollcommand=play_history_scroll_bar_v.set, xscrollcommand=play_history_scroll_bar_h.set)
    # 分隔线
    # separator = Frame(tk, height=2, bd=1, relief=SUNKEN)
    # separator.pack(fill="x", padx=2, pady=2)
    # 底部信息
    # 底部文字
    bottom_frm = Frame(tk)
    bottom_frm.pack(side="bottom")
    # 晨小明工作室
    cxm_frm = Frame(bottom_frm)
    cxm_frm.pack()
    original_image = Image.open(cxm_path)
    resized_image = original_image.resize((round(original_image.width / 21), round(original_image.height / 21)))  # 缩放图片到指定大小
    cxm_image_new = ImageTk.PhotoImage(resized_image)
    cv_cxm = Canvas(cxm_frm, width=110, height=cxm_image_new.height(), highlightthickness=0)
    cv_cxm.create_image(5, 0, image=cxm_image_new, anchor="nw")
    cv_cxm.grid(row=0, column=0)
    bottom_info_frm = Frame(bottom_frm)
    bottom_info_frm.pack()
    bottom_label_a = Label(bottom_info_frm, font=("Ya Hei", 10), fg="#888888", text="作者：晨小明")
    bottom_label_a.grid(row=0, column=0, padx=5, pady=5)
    bottom_label_v = Label(bottom_info_frm, font=("Ya Hei", 10), fg="#888888", text=f"版本：{VERSION}")
    bottom_label_v.grid(row=0, column=1, padx=5, pady=5)
    bottom_label_t = Label(bottom_info_frm, font=("Ya Hei", 10), fg="#888888", text=F"更新时间：{UPDATETIME}")
    bottom_label_t.grid(row=0, column=2, padx=5, pady=5)
    bottom_label_w = Label(bottom_info_frm, font=("Ya Hei", 10), fg="#888888", text="微信公众号：晨小明工作室（CXM-Studio）")
    bottom_label_w.grid(row=0, column=3, padx=5, pady=5)
    bottom_label_w.bind("<Enter>", on_enter)
    bottom_label_w.bind("<Leave>", on_leave)
    bottom_label_w.bind("<Button-1>", wxTk)
    bottom_label_c = Label(bottom_info_frm, font=("Ya Hei", 10), fg="#888888", text="联系作者：3038693133@qq.com")
    bottom_label_c.grid(row=0, column=4, padx=5, pady=5)
    bottom_label_c.bind("<Enter>", on_enter)
    bottom_label_c.bind("<Leave>", on_leave)
    bottom_label_c.bind("<Button-1>", toMail)
    tk.mainloop()
    # tkinter end
