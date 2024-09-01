import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# 创建主应用程序窗口
root = tk.Tk()
root.title("Excel Formula Generator")

# 标签和输入框
tk.Label(root, text="Start Column (e.g., c):").grid(row=0, column=0)
start_col_entry = tk.Entry(root)
start_col_entry.grid(row=0, column=1)

tk.Label(root, text="End Column (e.g., f):").grid(row=1, column=0)
end_col_entry = tk.Entry(root)
end_col_entry.grid(row=1, column=1)

tk.Label(root, text="Fixed Column (e.g., b):").grid(row=2, column=0)
fixed_col_entry = tk.Entry(root)
fixed_col_entry.grid(row=2, column=1)

# 生成Excel文件的函数
def generate_excel():
    start_col = start_col_entry.get().lower()
    end_col = end_col_entry.get().lower()
    fixed_col = fixed_col_entry.get().lower()
    
    if not start_col or not end_col or not fixed_col:
        messagebox.showerror("Error", "All fields must be filled out.")
        return

    start_index = col_to_num(start_col)
    end_index = col_to_num(end_col)
    fixed_index = col_to_num(fixed_col)

    if start_index >= end_index:
        messagebox.showerror("Error", "Start column must be before end column.")
        return
    
    data = {}
    for i in range(start_index, end_index + 1):
        col_letter = num_to_col(i)
        col_name = f"col({col_letter})-col({fixed_col})"
        data[col_letter] = [col_name]
    
    df = pd.DataFrame(data)
    
    # 选择保存文件的路径
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Excel file saved as {file_path}")

def col_to_num(col):
    """Convert Excel column letter to number."""
    num = 0
    for char in col:
        num = num * 26 + (ord(char) - ord('a') + 1)
    return num

def num_to_col(num):
    """Convert number to Excel column letter."""
    col = ""
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        col = chr(65 + remainder) + col
    return col.lower()

# 生成按钮
generate_button = tk.Button(root, text="Generate Excel", command=generate_excel)
generate_button.grid(row=3, columnspan=2)

# 运行主循环
root.mainloop()
