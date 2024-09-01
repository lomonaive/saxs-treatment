import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

class DataMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Merger")
        self.file_list = []

        # Setup UI
        self.frame = tk.Frame(root)
        self.frame.pack(padx=10, pady=10)

        self.btn_load = tk.Button(self.frame, text="Load Files", command=self.load_files)
        self.btn_load.pack(side=tk.TOP, fill=tk.X)

        self.lbl_files = tk.Label(self.frame, text="N∂o files selected")
        self.lbl_files.pack(side=tk.TOP, pady=(5, 0))

        self.file_listbox = tk.Listbox(self.frame, selectmode=tk.SINGLE)
        self.file_listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.btn_up = tk.Button(self.frame, text="Move Up", command=self.move_up)
        self.btn_up.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_down = tk.Button(self.frame, text="Move Down", command=self.move_down)
        self.btn_down.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_merge = tk.Button(self.frame, text="Merge and Save Excel", command=self.merge_files)
        self.btn_merge.pack(side=tk.BOTTOM, fill=tk.X)

    def load_files(self):
        filenames = filedialog.askopenfilenames(filetypes=[("Data files", "*.dat")])
        if filenames:
            self.file_list = list(filenames)
            self.update_file_listbox()
        else:
            self.lbl_files.config(text="No files selected")

    def update_file_listbox(self):
        self.file_listbox.delete(0, tk.END)
        for file in self.file_list:
            self.file_listbox.insert(tk.END, os.path.basename(file))
        self.lbl_files.config(text=f"{len(self.file_list)} files selected")

    def move_up(self):
        selected_index = self.file_listbox.curselection()
        if selected_index:
            index = selected_index[0]
            if index > 0:
                self.file_list[index], self.file_list[index - 1] = self.file_list[index - 1], self.file_list[index]
                self.update_file_listbox()
                self.file_listbox.selection_set(index - 1)

    def move_down(self):
        selected_index = self.file_listbox.curselection()
        if selected_index:
            index = selected_index[0]
            if index < len(self.file_list) - 1:
                self.file_list[index], self.file_list[index + 1] = self.file_list[index + 1], self.file_list[index]
                self.update_file_listbox()
                self.file_listbox.selection_set(index + 1)

    def merge_files(self):
        if not self.file_list:
            messagebox.showerror("Error", "No files loaded to merge!")
            return

        combined_df = pd.DataFrame()
        col_offset = 0

        for file in self.file_list:
            try:
                with open(file, 'r') as f:
                    # 读取文件的前几行来提取备注信息
                    headers = [next(f) for _ in range(5)]  # 假设前5行包含有用的头部信息

                # 使用pandas读取数据，跳过头部
                df = pd.read_csv(file, delim_whitespace=True, skiprows=5, header=None)
                num_cols = df.shape[1]
                header_labels = ['{}_{}'.format(headers[0].strip(), i) for i in range(num_cols)]
                df.columns = header_labels
                # 重新索引列以避免重叠
                new_columns = {df.columns[i]: df.columns[i] + str(col_offset) for i in range(len(df.columns))}
                df.rename(columns=new_columns, inplace=True)
                combined_df = pd.concat([combined_df, df], axis=1)
                col_offset += num_cols
            except Exception as e:
                messagebox.showerror("Error", f"Failed to parse {file}: {e}")
                return

        save_path = filedialog.asksaveasfilename(filetypes=[("Excel files", "*.xlsx")], defaultextension=".xlsx")
        if save_path:
            combined_df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Data merged and saved to {save_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataMergerApp(root)
    root.mainloop()
