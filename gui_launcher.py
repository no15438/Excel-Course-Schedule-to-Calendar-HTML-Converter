import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from xlsx_to_calendar import process_excel_file, generate_course_calendar, generate_ics_calendar, get_term_dates

def generate_calendars():
    # 验证选择
    if not selected_file.get():
        messagebox.showwarning("警告", "请先选择一个 Excel 文件。")
        return
    if not output_dir.get():
        messagebox.showwarning("警告", "请先选择一个导出目录。")
        return
    
    # 启动进度条（不确定模式）
    progress_bar.grid(row=4, column=0, columnspan=3, pady=10)
    progress_bar.start()
    root.update_idletasks()
    
    selected_term = term_var.get()  # 'term1' or 'term2'
    try:
        term_start, term_end = get_term_dates(selected_term)
        df = process_excel_file(selected_file.get())
        
        # 在导出目录生成HTML日历
        html_content = generate_course_calendar(df, selected_term)
        html_path = os.path.join(output_dir.get(), 'course_calendar.html')
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        # 在导出目录生成ICS日历
        cal = generate_ics_calendar(df, term_start, term_end, selected_term)
        ics_path = os.path.join(output_dir.get(), 'course_calendar.ics')
        with open(ics_path, 'wb') as f:
            f.write(cal.to_ical())
        
        progress_bar.stop()
        progress_bar.grid_remove()
        
        messagebox.showinfo("成功", f"HTML 与 ICS 日历已成功生成！\n请查看目录：\n{output_dir.get()}")
    except Exception as e:
        progress_bar.stop()
        progress_bar.grid_remove()
        messagebox.showerror("错误", f"生成失败：{str(e)}")

def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx *.xls")])
    if file_path:
        selected_file.set(file_path)

def choose_directory():
    directory = filedialog.askdirectory()
    if directory:
        output_dir.set(directory)

root = tk.Tk()
root.title("UBC Workday Excel课表转换器")

selected_file = tk.StringVar()
output_dir = tk.StringVar()
term_var = tk.StringVar(value='term1')

# 文件选择区
file_frame = tk.Frame(root)
file_frame.grid(row=0, column=0, columnspan=3, pady=10, padx=10, sticky='w')
tk.Label(file_frame, text="选择Excel文件:").grid(row=0, column=0, padx=5, sticky='e')
tk.Entry(file_frame, textvariable=selected_file, width=50).grid(row=0, column=1, padx=5)
tk.Button(file_frame, text="浏览...", command=choose_file).grid(row=0, column=2, padx=5)

# 导出目录选择区
out_frame = tk.Frame(root)
out_frame.grid(row=1, column=0, columnspan=3, pady=10, padx=10, sticky='w')
tk.Label(out_frame, text="选择导出目录:").grid(row=0, column=0, padx=5, sticky='e')
tk.Entry(out_frame, textvariable=output_dir, width=50).grid(row=0, column=1, padx=5)
tk.Button(out_frame, text="浏览...", command=choose_directory).grid(row=0, column=2, padx=5)

# 学期选择区
term_frame = tk.Frame(root)
term_frame.grid(row=2, column=0, columnspan=3, pady=10, padx=10, sticky='w')
tk.Label(term_frame, text="选择学期：").grid(row=0, column=0, padx=5)
tk.Radiobutton(term_frame, text="Term 1 (2024/09/03-2024/12/06)", variable=term_var, value='term1').grid(row=0, column=1, padx=5)
tk.Radiobutton(term_frame, text="Term 2 (2025/01/06-2025/04/06)", variable=term_var, value='term2').grid(row=0, column=2, padx=5)

# 生成按钮
tk.Button(root, text="生成日历", command=generate_calendars, bg="#4CAF50", fg="white", padx=20, pady=5).grid(row=3, column=0, columnspan=3, pady=20)

# 进度条（初始隐藏）
progress_bar = ttk.Progressbar(root, orient='horizontal', mode='indeterminate', length=300)

root.resizable(False, False)
root.mainloop()