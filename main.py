from pdf2docx import Converter
import maliang
import tkinter.messagebox as messagebox
from tkinter import filedialog  # 导入filedialog模块

# 创建窗口
root = maliang.Tk(title="PDF 转 Word")
root.iconbitmap('output.ico')
root.center()

# 使用一个标志位确保只执行一次转换
convert_in_progress = False

canv = maliang.Canvas(auto_zoom=True, keep_ratio="min", free_anchor=True)
canv.place(width=1280, height=720, x=640, y=360, anchor="center")
maliang.Text(canv, (640, 200), text="PDF 转 Word", fontsize=48, anchor="center")

# 添加控件
maliang.Text(canv, (450, 300), text="pdf文件", anchor="nw")
pdf_input = maliang.InputBox(canv, (450, 340), (380, 50), placeholder="点击输入pdf文件路径")

maliang.Button(canv, (840, 340), (180, 50), text="选择文件", anchor="nw", fontsize=24, command=lambda: select_file(pdf_input))
maliang.Button(canv, (450, 400), (240, 50), text="转换文件", anchor="nw", fontsize=24, command=lambda: convert_file(pdf_input))
path_word = maliang.Text(canv, (450, 500), text="转换word文件位置:", anchor="nw",)

# 选择文件的函数
def select_file(pdf_input):
    # 打开文件选择框，获取文件路径
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        # 将选择的文件路径设置到输入框
        pdf_input.set(file_path)
        path_word.set("转换word文件位置: " + file_path.replace('.pdf', '.docx'))

# 转换文件的函数
def convert_file(input_box):
    global convert_in_progress  # 使用全局变量来跟踪是否正在转换
    
    if convert_in_progress:
        messagebox.showwarning("警告", "文件正在转换中，请稍后再试。")
        return
    
    pdf_file = input_box.get()  # 获取输入框中的文件路径
    if pdf_file:
        docx_file = pdf_file.replace('.pdf', '.docx')  # 默认保存路径为同名的docx文件
        try:
            convert_in_progress = True  # 设置标志位，表示正在转换

            # 调用pdf2docx转换
            cv = Converter(pdf_file)
            cv.convert(docx_file, multi_processing=False)  # 转换所有页面
            cv.close()
            messagebox.showinfo("成功", "转换成功！文件保存到: " + docx_file)
            
        except Exception as e:
            # 如果有错误发生，弹出一次错误消息框
            messagebox.showerror("错误", f"转换失败: {str(e)}")
        
        finally:
            convert_in_progress = False  # 重置标志位，表示转换完成

# 关闭窗口时的处理函数

root.mainloop()
