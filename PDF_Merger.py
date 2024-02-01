import os
import fitz
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from tkinter.simpledialog import askstring

class PDFMergerGUI:
    def __init__(self, window):
        self.window = window
        self.window.title("PDF合并工具")

        # 创建文件夹路径选择部件
        folder_path_label = tk.Label(window, text="选择包含PDF文件的文件夹:")
        folder_path_label.pack(padx=10, pady=10, anchor=tk.W)
        self.folder_path_entry = tk.Entry(window, width=50)
        self.folder_path_entry.pack(padx=10, pady=10)
        folder_path_button = tk.Button(window, text="浏览", command=self.select_folder)
        folder_path_button.pack(padx=10, pady=10)

        # 创建输出文件选择部件
        output_file_label = tk.Label(window, text="选择合并后的PDF文件保存的位置:")
        output_file_label.pack(padx=10, pady=10, anchor=tk.W)
        self.output_file_entry = tk.Entry(window, width=50)
        self.output_file_entry.pack(padx=10, pady=10)
        output_file_button = tk.Button(window, text="浏览", command=self.select_output_file)
        output_file_button.pack(padx=10, pady=10)

        # 创建页数选择部件
        self.page_range_label = tk.Label(window, text="输入要合并的页码范围(例如1-3):")
        self.page_range_label.pack(padx=10, pady=10, anchor=tk.W)
        self.page_range_entry = tk.Entry(window, width=10)
        self.page_range_entry.pack(padx=10, pady=10)
        self.page_range_button = tk.Button(window, text="选择页码范围", command=self.select_page_range, state=tk.DISABLED)
        self.page_range_button.pack(padx=10, pady=10)

        # 创建设置方法选择部件
        self.method_var = tk.StringVar()
        self.method_var.set("A")
        method_label = tk.Label(window, text="设置方法:")
        method_label.pack(padx=10, pady=10, anchor=tk.W)
        method_radio_a = tk.Radiobutton(window, text="选取所有的PDF第a页-第b页", variable=self.method_var, value="A", command=self.update_page_range_button_state)
        method_radio_a.pack(padx=10, pady=10, anchor=tk.W)
        method_radio_b = tk.Radiobutton(window, text="针对每个PDF分别设置选取第a页-第b页", variable=self.method_var, value="B", command=self.update_page_range_button_state)
        method_radio_b.pack(padx=10, pady=10, anchor=tk.W)

        # 创建预览按钮
        preview_button = tk.Button(window, text="预览选页", command=self.preview_pages)
        preview_button.pack(padx=10, pady=10)

        # 创建合并按钮
        merge_button = tk.Button(window, text="合并PDF", command=self.merge_pdf)
        merge_button.pack(padx=10, pady=10)

        # 创建状态标签
        self.status_label = tk.Label(window, text="")
        self.status_label.pack(padx=10, pady=10)

        # 创建预览图像区域
        self.preview_window = None
        self.preview_frame = None
        self.prev_button = None
        self.next_button = None
        self.preview_canvas = None
        self.pdf_listbox = None
        self.page_range_listbox = None
        self.preview_images = []
        self.current_page = 0

        # 创建滑动条
        self.image_quality_scale = tk.Scale(window, from_=100, to=2000, orient=tk.HORIZONTAL, length=200, resolution=100, label="预览清晰度")
        self.image_quality_scale.set(800)
        self.image_quality_scale.pack(padx=10, pady=10)

        # Make components resizable with window size
        self.window.rowconfigure(0, weight=1)
        self.window.columnconfigure(0, weight=1)

        # Create preview frame
        self.preview_frame = tk.Frame(self.window)
        self.preview_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path_entry.delete(0, tk.END)
            self.folder_path_entry.insert(0, folder_path)
            self.output_file_entry.delete(0, tk.END)
            self.output_file_entry.insert(0, os.path.join(folder_path, "merged.pdf"))

    def select_output_file(self):
        output_file = filedialog.asksaveasfilename(defaultextension=".pdf")
        if output_file:
            self.output_file_entry.delete(0, tk.END)
            self.output_file_entry.insert(0, output_file)

    def select_page_range(self):
        page_range = askstring("选择页码范围", "输入要合并的页码范围(例如1-3):")
        if page_range:
            self.page_range_entry.delete(0, tk.END)
            self.page_range_entry.insert(0, page_range)

    def update_page_range_button_state(self, *args):
        if self.method_var.get() == "B":
            self.page_range_button.config(state=tk.NORMAL)
        else:
            self.page_range_button.config(state=tk.DISABLED)

    def preview_pages(self):
        folder_path = self.folder_path_entry.get()
        page_range = self.page_range_entry.get()
        image_quality = self.image_quality_scale.get()

        if not folder_path or not page_range:
            messagebox.showwarning("警告", "请先选择文件夹路径和设置页数范围")
            return

        pdf_files = sorted([filename for filename in os.listdir(folder_path) if filename.lower().endswith('.pdf')], key=lambda x: int(''.join(filter(str.isdigit, x))) if any(char.isdigit() for char in x) else x)
        if not pdf_files:
            messagebox.showwarning("警告", "文件夹中没有PDF文件")
            return

        self.preview_images = []
        self.current_page = 0

        # 清空预览图像区域和PDF文件选择列表框
        if self.preview_window:
            self.preview_window.destroy()

        # 创建预览选页窗口
        self.preview_window = tk.Toplevel(self.window)
        self.preview_window.title("预览选页")
        self.preview_frame = tk.Frame(self.preview_window)
        self.preview_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # 创建滚动条
        scrollbar = tk.Scrollbar(self.preview_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 创建预览图像区域
        self.preview_canvas = tk.Canvas(self.preview_frame, yscrollcommand=scrollbar.set)
        self.preview_canvas.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.BOTH, expand=True)

        # 配置滚动条与预览图像区域的关联
        scrollbar.config(command=self.preview_canvas.yview)

        # 创建PDF文件选择列表框
        self.pdf_listbox = tk.Listbox(self.preview_frame, width=50)
        self.pdf_listbox.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.Y)

        # 创建页码范围列表框
        self.page_range_listbox = tk.Listbox(self.preview_frame, width=10)
        self.page_range_listbox.pack(side=tk.LEFT, padx=5, pady=5, fill=tk.Y)

        # 创建切换按钮
        button_frame = tk.Frame(self.preview_window)
        button_frame.pack(side=tk.BOTTOM, padx=10, pady=10)
        self.prev_button = tk.Button(button_frame, text="上一页", command=self.show_prev_page, state=tk.DISABLED)
        self.prev_button.pack(side=tk.LEFT, padx=10, pady=10)
        self.next_button = tk.Button(button_frame, text="下一页", command=self.show_next_page, state=tk.DISABLED)
        self.next_button.pack(side=tk.LEFT, padx=10, pady=10)

        # 绑定鼠标滚轮的事件
        self.preview_window.bind("<MouseWheel>", self.scroll_preview)

        # 遍历PDF文件并加载预览图像
        for filename in pdf_files:
            file_path = os.path.join(folder_path, filename)
            pdf_reader = fitz.open(file_path)

            page_range_parts = page_range.split('-')
            start_page = int(page_range_parts[0]) - 1
            end_page = int(page_range_parts[-1]) - 1 if len(page_range_parts) > 1 else int(page_range_parts[0]) - 1

            for page_num in range(start_page, min(end_page+1, pdf_reader.page_count)):
                page = pdf_reader.load_page(page_num)
                pix = page.get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                img.thumbnail((image_quality, image_quality))
                img_tk = ImageTk.PhotoImage(img)

                self.preview_images.append(img_tk)

            self.pdf_listbox.insert(tk.END, filename)
            self.page_range_listbox.insert(tk.END, page_range)

            pdf_reader.close()

        if self.preview_images:
            self.show_page(0)
            self.prev_button.config(state=tk.DISABLED)
            self.next_button.config(state=tk.NORMAL)
        else:
            messagebox.showwarning("警告", "没有找到符合条件的页码范围")

    def show_page(self, page_num):
        img_tk = self.preview_images[page_num]
        self.preview_canvas.delete("all")
        self.preview_canvas.create_image(0, 0, anchor=tk.NW, image=img_tk)
        self.preview_canvas.config(scrollregion=self.preview_canvas.bbox("all"))

    def show_prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.show_page(self.current_page)
            self.next_button.config(state=tk.NORMAL)
        if self.current_page == 0:
            self.prev_button.config(state=tk.DISABLED)

    def show_next_page(self):
        if self.current_page < len(self.preview_images) - 1:
            self.current_page += 1
            self.show_page(self.current_page)
            self.prev_button.config(state=tk.NORMAL)
        if self.current_page == len(self.preview_images) - 1:
            self.next_button.config(state=tk.DISABLED)

    def update_preview(self, event):
        selected_index = self.pdf_listbox.curselection()
        if selected_index:
            selected_index = int(selected_index[0])
            self.show_page(selected_index)

    def scroll_preview(self, event):
        if event.delta > 0 and not self.pdf_listbox.bbox(tk.ACTIVE):
            self.show_prev_page()
        elif event.delta < 0 and not self.pdf_listbox.bbox(tk.ACTIVE):
            self.show_next_page()

    def merge_pdf(self):
        folder_path = self.folder_path_entry.get()
        output_file = self.output_file_entry.get()
        method = self.method_var.get()

        if not folder_path or not output_file:
            messagebox.showwarning("警告", "请先选择文件夹路径和输出文件")
            return

        pdf_files = sorted([filename for filename in os.listdir(folder_path) if filename.lower().endswith('.pdf')], key=lambda x: int(''.join(filter(str.isdigit, x))) if any(char.isdigit() for char in x) else x)
        if not pdf_files:
            messagebox.showwarning("警告", "文件夹中没有PDF文件")
            return

        pdf_writer = fitz.open()

        progress_bar = ttk.Progressbar(self.window, orient=tk.HORIZONTAL, length=200, mode='determinate')
        progress_bar.pack(padx=10, pady=10)

        for i, filename in enumerate(pdf_files):
            page_range = self.page_range_listbox.get(i)
            file_path = os.path.join(folder_path, filename)
            pdf_reader = fitz.open(file_path)

            page_range_parts = page_range.split('-')
            start_page = int(page_range_parts[0]) - 1
            end_page = int(page_range_parts[-1]) - 1 if len(page_range_parts) > 1 else int(page_range_parts[0]) - 1

            if method == "A":
                pdf_writer.insert_pdf(pdf_reader, from_page=start_page, to_page=min(end_page, pdf_reader.page_count-1))
            elif method == "B":
                page_range = askstring("选择页码范围", f"在模式B中分别选取每个pdf不同的{filename}页，输入要合并的页码范围(例如1-3):")
                if page_range:
                    page_range_parts = page_range.split('-')
                    start_page = int(page_range_parts[0]) - 1
                    end_page = int(page_range_parts[-1]) - 1 if len(page_range_parts) > 1 else int(page_range_parts[0]) - 1
                    pdf_writer.insert_pdf(pdf_reader, from_page=start_page, to_page=end_page)

            pdf_reader.close()

            progress_value = (i + 1) / len(pdf_files) * 100
            progress_bar['value'] = progress_value
            self.window.update()

        pdf_writer.save(output_file)
        pdf_writer.close()

        self.status_label.config(text="PDF合并完成！")

        messagebox.showinfo("提示", "PDF合并完成！")

        progress_bar.destroy()

# 创建主窗口
window = tk.Tk()
pdf_merger_gui = PDFMergerGUI(window)

# 运行主循环
try:
    window.mainloop()
except Exception as e:
    print("出现错误:", e)