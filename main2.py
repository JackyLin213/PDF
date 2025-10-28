import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pypdf import PdfReader, PdfWriter # 確保沒有 import PdfMerger
import os

class PdfToolkitApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Editor (Delete, Merge, Inesert)")
        self.root.geometry("600x500")  # 再次加大視窗

        # --- 建立分頁控制器 ---
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(pady=10, padx=10, fill="both", expand=True)

        # --- 分頁 1: 刪除頁面 ---
        self.delete_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(self.delete_frame, text='刪除頁面')
        self.create_delete_tab(self.delete_frame)

        # --- 分頁 2: 合併 PDF (A+B+C) ---
        self.merge_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(self.merge_frame, text='合併 PDF (依序)')
        self.create_merge_tab(self.merge_frame)
        
        # --- 分頁 3: 插入 PDF (A + B + A) ---
        self.insert_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(self.insert_frame, text='插入 PDF (指定位置)')
        self.create_insert_tab(self.insert_frame)

    # =====================================================================
    # 輔助函式: 選擇單一 PDF 檔案
    # =====================================================================
    def select_pdf_file(self, string_var_to_update):
        """開啟檔案對話框，並將選擇的路徑更新到指定的 StringVar"""
        file_path = filedialog.askopenfilename(
            title="請選擇一個 PDF 檔案",
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")]
        )
        if file_path:
            string_var_to_update.set(file_path)

    # =====================================================================
    # 分頁 1: 刪除頁面 (Delete Tab)
    # =====================================================================
    def create_delete_tab(self, parent_frame):
        self.delete_input_path = tk.StringVar()

        input_frame = ttk.Frame(parent_frame)
        input_frame.pack(fill="x", pady=10)
        ttk.Label(input_frame, text="來源 PDF:").pack(side="left", padx=5)
        self.delete_input_entry = ttk.Entry(input_frame, textvariable=self.delete_input_path, width=40, state="readonly")
        self.delete_input_entry.pack(side="left", fill="x", expand=True, padx=5)
        # 使用 lambda 將參數傳遞給輔助函式
        ttk.Button(input_frame, text="瀏覽...", command=lambda: self.select_pdf_file(self.delete_input_path)).pack(side="left", padx=5)

        pages_frame = ttk.Frame(parent_frame)
        pages_frame.pack(fill="x", pady=10)
        ttk.Label(pages_frame, text="刪除頁碼:").pack(side="left", padx=5)
        self.delete_pages_entry = ttk.Entry(pages_frame, width=40)
        self.delete_pages_entry.pack(side="left", fill="x", expand=True, padx=5)

        ttk.Label(parent_frame, text="請輸入頁碼（從 1 開始），用逗號分隔，或使用連字號指定範圍。\n範例: 1, 3, 5-8", foreground="gray").pack(fill="x", padx=10)
        ttk.Button(parent_frame, text="刪除頁面並另存新檔...", command=self.process_delete_and_save).pack(pady=20, fill="x")

    def parse_page_string(self, page_str: str) -> list[int]:
        pages_to_delete_indices = set()
        if not page_str.strip():
            return []
        parts = page_str.split(',')
        for part in parts:
            part = part.strip()
            if not part: continue
            if '-' in part:
                start_str, end_str = part.split('-')
                start = int(start_str.strip())
                end = int(end_str.strip())
                if start > end: start, end = end, start
                for i in range(start, end + 1):
                    pages_to_delete_indices.add(i - 1)
            else:
                pages_to_delete_indices.add(int(part.strip()) - 1)
        return sorted(list(pages_to_delete_indices))

    def process_delete_and_save(self):
        input_path = self.delete_input_path.get()
        pages_str = self.delete_pages_entry.get()

        if not input_path:
            messagebox.showerror("錯誤", "請先選擇一個 PDF 檔案！", parent=self.delete_frame)
            return
        if not pages_str:
            messagebox.showerror("錯誤", "請輸入要刪除的頁碼！", parent=self.delete_frame)
            return

        try:
            pages_to_delete = self.parse_page_string(pages_str)
        except Exception:
            messagebox.showerror("錯誤", "頁碼格式錯誤！", parent=self.delete_frame)
            return
        if not pages_to_delete:
            messagebox.showwarning("提示", "您沒有輸入任何有效的頁碼。", parent=self.delete_frame)
            return

        default_name = os.path.basename(input_path).replace(".pdf", "_deleted.pdf")
        output_path = filedialog.asksaveasfilename(title="請選擇儲存位置", initialfile=default_name, defaultextension=".pdf", filetypes=[("PDF 檔案", "*.pdf")])
        if not output_path: return

        try:
            reader = PdfReader(input_path)
            writer = PdfWriter()
            total_pages = len(reader.pages)
            invalid_pages = [p + 1 for p in pages_to_delete if p >= total_pages]
            if invalid_pages:
                messagebox.showerror("錯誤", f"頁碼 {invalid_pages} 超出檔案總頁數 ({total_pages})！", parent=self.delete_frame)
                return

            for i in range(total_pages):
                if i not in pages_to_delete:
                    writer.add_page(reader.pages[i])

            with open(output_path, "wb") as f:
                writer.write(f)
            messagebox.showinfo("成功", f"已成功刪除頁面！\n新檔案儲存於:\n{output_path}", parent=self.delete_frame)
        except Exception as e:
            messagebox.showerror("處理失敗", f"讀取或寫入 PDF 時發生錯誤:\n{e}", parent=self.delete_frame)

    # =====================================================================
    # 分頁 2: 合併 PDF (Merge Tab - A+B+C)
    # =====================================================================
    def create_merge_tab(self, parent_frame):
        list_frame = ttk.Frame(parent_frame)
        list_frame.pack(fill="both", expand=True, pady=10)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
        self.merge_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, selectmode="extended") # 允許多選
        scrollbar.config(command=self.merge_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.merge_listbox.pack(side="left", fill="both", expand=True)

        order_frame = ttk.Frame(parent_frame)
        order_frame.pack(fill="x")
        ttk.Button(order_frame, text="上移", command=self.move_up).pack(side="left", padx=5, pady=5)
        ttk.Button(order_frame, text="下移", command=self.move_down).pack(side="left", padx=5, pady=5)

        button_frame = ttk.Frame(parent_frame)
        button_frame.pack(fill="x")
        ttk.Button(button_frame, text="新增檔案...", command=self.add_merge_files).pack(side="left", padx=5, pady=5)
        ttk.Button(button_frame, text="移除選定檔案", command=self.remove_merge_file).pack(side="left", padx=5, pady=5)
        ttk.Button(button_frame, text="全部清除", command=self.clear_merge_list).pack(side="left", padx=5, pady=5)

        ttk.Button(parent_frame, text="合併並另存新檔...", command=self.process_merge_and_save).pack(fill="x", pady=20)

    def add_merge_files(self):
        file_paths = filedialog.askopenfilenames(title="請選擇要合併的 PDF 檔案", filetypes=[("PDF 檔案", "*.pdf")])
        for file_path in file_paths:
            self.merge_listbox.insert("end", file_path)

    def remove_merge_file(self):
        selected_indices = self.merge_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("提示", "請先在列表中選擇一個要移除的檔案。", parent=self.merge_frame)
            return
        for index in reversed(selected_indices):
            self.merge_listbox.delete(index)

    def clear_merge_list(self):
        self.merge_listbox.delete(0, "end")

    def move_up(self):
        selected_indices = self.merge_listbox.curselection()
        if not selected_indices: return
        index = selected_indices[0] # 只處理單選
        if index == 0: return
        file_path = self.merge_listbox.get(index)
        self.merge_listbox.delete(index)
        self.merge_listbox.insert(index - 1, file_path)
        self.merge_listbox.selection_set(index - 1)

    def move_down(self):
        selected_indices = self.merge_listbox.curselection()
        if not selected_indices: return
        index = selected_indices[0] # 只處理單選
        if index == self.merge_listbox.size() - 1: return
        file_path = self.merge_listbox.get(index)
        self.merge_listbox.delete(index)
        self.merge_listbox.insert(index + 1, file_path)
        self.merge_listbox.selection_set(index + 1)

    def process_merge_and_save(self):
        # (使用方案 2: PdfWriter)
        file_paths = self.merge_listbox.get(0, "end")
        if len(file_paths) < 2:
            messagebox.showerror("錯誤", "請至少新增兩個 PDF 檔案才能合併。", parent=self.merge_frame)
            return

        output_path = filedialog.asksaveasfilename(title="請選擇合併後的儲存位置", initialfile="merged_document.pdf", defaultextension=".pdf", filetypes=[("PDF 檔案", "*.pdf")])
        if not output_path: return

        try:
            writer = PdfWriter()
            for pdf_path in file_paths:
                writer.append(pdf_path) # 依序附加
            with open(output_path, "wb") as f:
                writer.write(f)
            messagebox.showinfo("成功", f"已成功合併 {len(file_paths)} 個檔案！\n新檔案儲存於:\n{output_path}", parent=self.merge_frame)
        except Exception as e:
            messagebox.showerror("合併失敗", f"處理 PDF 時發生錯誤:\n{e}", parent=self.merge_frame)

    # =====================================================================
    # 分頁 3: 插入 PDF (Insert Tab)
    # =====================================================================
    def create_insert_tab(self, parent_frame):
        # --- 變數 ---
        self.insert_base_path = tk.StringVar()
        self.insert_new_path = tk.StringVar()

        # --- 1. 選擇主要檔案 (Base PDF) ---
        base_frame = ttk.Frame(parent_frame)
        base_frame.pack(fill="x", pady=10)
        ttk.Label(base_frame, text="主要 PDF:").pack(side="left", padx=5)
        self.insert_base_entry = ttk.Entry(base_frame, textvariable=self.insert_base_path, width=40, state="readonly")
        self.insert_base_entry.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(base_frame, text="瀏覽...", command=lambda: self.select_pdf_file(self.insert_base_path)).pack(side="left", padx=5)

        # --- 2. 選擇要插入的檔案 (Insert PDF) ---
        new_frame = ttk.Frame(parent_frame)
        new_frame.pack(fill="x", pady=10)
        ttk.Label(new_frame, text="要插入的 PDF:").pack(side="left", padx=5)
        self.insert_new_entry = ttk.Entry(new_frame, textvariable=self.insert_new_path, width=40, state="readonly")
        self.insert_new_entry.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(new_frame, text="瀏覽...", command=lambda: self.select_pdf_file(self.insert_new_path)).pack(side="left", padx=5)

        # --- 3. 輸入要插入的頁碼 ---
        pages_frame = ttk.Frame(parent_frame)
        pages_frame.pack(fill="x", pady=10)
        ttk.Label(pages_frame, text="插入到第 N 頁:").pack(side="left", padx=5)
        self.insert_page_entry = ttk.Entry(pages_frame, width=40)
        self.insert_page_entry.pack(side="left", fill="x", expand=True, padx=5)

        # --- 4. 頁碼說明 ---
        hint_label = ttk.Label(parent_frame, text="請輸入一個頁碼（從 1 開始）。\n範例: 輸入 1，會插入到檔案最前面。\n範例: 輸入 5，會插入到原第 4 頁之後（新內容從第 5 頁開始）。", foreground="gray")
        hint_label.pack(fill="x", padx=10)

        # --- 5. 執行按鈕 ---
        process_button = ttk.Button(parent_frame, text="插入並另存新檔...", command=self.process_insert_and_save)
        process_button.pack(pady=20, fill="x")

    def process_insert_and_save(self):
        """核心插入邏輯 (不需 merge() 或 append())"""
        base_path = self.insert_base_path.get()
        insert_path = self.insert_new_path.get()
        page_str = self.insert_page_entry.get()

        # --- 1. 驗證輸入 ---
        if not base_path:
            messagebox.showerror("錯誤", "請先選擇一個「主要」 PDF 檔案！", parent=self.insert_frame)
            return
        if not insert_path:
            messagebox.showerror("錯誤", "請先選擇一個「要插入的」 PDF 檔案！", parent=self.insert_frame)
            return
        if not page_str:
            messagebox.showerror("錯誤", "請輸入要插入的頁碼！", parent=self.insert_frame)
            return

        try:
            page_num_to_insert = int(page_str)
        except ValueError:
            messagebox.showerror("錯誤", "頁碼必須是數字！", parent=self.insert_frame)
            return

        if page_num_to_insert < 1:
            messagebox.showerror("錯誤", "頁碼必須大於等於 1！", parent=self.insert_frame)
            return

        # --- 2. 詢問儲存路徑 ---
        default_name = os.path.basename(base_path).replace(".pdf", "_inserted.pdf")
        output_path = filedialog.asksaveasfilename(
            title="請選擇儲存位置",
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF 檔案", "*.pdf")]
        )
        if not output_path:
            return # 使用者取消

        # --- 3. 執行 pypdf 插入邏輯 (手動) ---
        try:
            writer = PdfWriter()
            base_reader = PdfReader(base_path)
            insert_reader = PdfReader(insert_path)
            
            total_base_pages = len(base_reader.pages)
            
            # 轉換為 0-based 索引
            insert_at_index = page_num_to_insert - 1

            # 驗證頁碼範圍
            # 允許插在最後一頁之後 (total_base_pages + 1)，索引即為 total_base_pages
            if insert_at_index > total_base_pages:
                messagebox.showerror("錯誤", f"頁碼 {page_num_to_insert} 超出範圍！\n此檔案共 {total_base_pages} 頁，您最多只能插入在第 {total_base_pages + 1} 頁 (即結尾)。", parent=self.insert_frame)
                return

            # 迴圈 1: 加入「插入點」之前的頁面
            for i in range(insert_at_index):
                writer.add_page(base_reader.pages[i])

            # 迴圈 2: 加入「要插入的 PDF」的所有頁面
            for page in insert_reader.pages:
                writer.add_page(page)

            # 迴圈 3: 加入「插入點」之後的剩餘頁面
            for i in range(insert_at_index, total_base_pages):
                writer.add_page(base_reader.pages[i])

            # 寫入新檔案
            with open(output_path, "wb") as f:
                writer.write(f)

            messagebox.showinfo("成功", f"已成功將檔案插入到第 {page_num_to_insert} 頁！\n新檔案儲存於:\n{output_path}", parent=self.insert_frame)
            
        except Exception as e:
            messagebox.showerror("處理失敗", f"讀取或寫入 PDF 時發生錯誤:\n{e}", parent=self.insert_frame)


if __name__ == "__main__":
    root = tk.Tk()
    app = PdfToolkitApp(root)
    root.mainloop()