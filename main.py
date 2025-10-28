import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pypdf import PdfReader, PdfWriter
import os

class PdfDeleterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 頁面刪除工具 (by Gemini)")
        self.root.geometry("500x250")  # 設定視窗大小

        # --- 變數 ---
        self.input_file_path = tk.StringVar()

        # --- 介面框架 ---
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill="both", expand=True)

        # --- 1. 選擇檔案 ---
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill="x", pady=10)

        input_label = ttk.Label(input_frame, text="來源 PDF:")
        input_label.pack(side="left", padx=5)

        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_file_path, width=40, state="readonly")
        self.input_entry.pack(side="left", fill="x", expand=True, padx=5)

        browse_button = ttk.Button(input_frame, text="瀏覽...", command=self.select_input_file)
        browse_button.pack(side="left", padx=5)

        # --- 2. 輸入要刪除的頁面 ---
        pages_frame = ttk.Frame(main_frame)
        pages_frame.pack(fill="x", pady=10)

        pages_label = ttk.Label(pages_frame, text="刪除頁碼:")
        pages_label.pack(side="left", padx=5)

        self.pages_entry = ttk.Entry(pages_frame, width=40)
        self.pages_entry.pack(side="left", fill="x", expand=True, padx=5)

        # --- 3. 頁碼說明 ---
        hint_label = ttk.Label(main_frame, text="請輸入頁碼（從 1 開始），用逗號分隔，或使用連字號指定範圍。\n範例: 1, 3, 5-8", foreground="gray")
        hint_label.pack(fill="x", padx=10)

        # --- 4. 執行按鈕 ---
        process_button = ttk.Button(main_frame, text="刪除頁面並另存新檔...", command=self.process_and_save)
        process_button.pack(pady=20, fill="x")

    def select_input_file(self):
        """開啟檔案對話框來選擇 PDF 檔案"""
        file_path = filedialog.askopenfilename(
            title="請選擇一個 PDF 檔案",
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")]
        )
        if file_path:
            self.input_file_path.set(file_path)
            
    def parse_page_string(self, page_str: str) -> list[int]:
        """
        將 "1, 3, 5-8" 這樣的字串轉換為 [0, 2, 4, 5, 6, 7] (0-based 索引)
        """
        pages_to_delete_indices = set()
        if not page_str.strip():
            return []
            
        parts = page_str.split(',')
        for part in parts:
            part = part.strip()
            if not part:
                continue
            
            if '-' in part:
                # 處理範圍 (e.g., "5-8")
                start_str, end_str = part.split('-')
                start = int(start_str)
                end = int(end_str)
                if start > end:
                    start, end = end, start # 交換
                # 轉換為 0-based 索引
                for i in range(start, end + 1):
                    pages_to_delete_indices.add(i - 1)
            else:
                # 處理單一頁面 (e.g., "3")
                page = int(part)
                # 轉換為 0-based 索引
                pages_to_delete_indices.add(page - 1)
                
        return sorted(list(pages_to_delete_indices))

    def process_and_save(self):
        """核心處理邏輯：讀取、刪除、儲存"""
        input_path = self.input_file_path.get()
        pages_str = self.pages_entry.get()

        # --- 1. 驗證輸入 ---
        if not input_path:
            messagebox.showerror("錯誤", "請先選擇一個 PDF 檔案！")
            return
            
        if not pages_str:
            messagebox.showerror("錯誤", "請輸入要刪除的頁碼！")
            return

        try:
            pages_to_delete = self.parse_page_string(pages_str)
        except ValueError:
            messagebox.showerror("錯誤", "頁碼格式錯誤！請使用數字、逗號和連字號。")
            return
        except Exception as e:
            messagebox.showerror("錯誤", f"頁碼解析失敗: {e}")
            return

        if not pages_to_delete:
            messagebox.showwarning("提示", "您沒有輸入任何有效的頁碼。")
            return

        # --- 2. 詢問儲存路徑 ---
        # 建議一個預設的檔名
        default_name = os.path.basename(input_path).replace(".pdf", "_modified.pdf")
        output_path = filedialog.asksaveasfilename(
            title="請選擇儲存位置",
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF 檔案", "*.pdf")]
        )

        if not output_path:
            # 使用者取消了儲存
            return

        # --- 3. 執行 pypdf 刪除邏輯 ---
        try:
            reader = PdfReader(input_path)
            writer = PdfWriter()
            
            total_pages = len(reader.pages)
            
            # 檢查頁碼是否超出範圍
            invalid_pages = [p + 1 for p in pages_to_delete if p >= total_pages]
            if invalid_pages:
                messagebox.showerror("錯誤", f"頁碼 {invalid_pages} 超出檔案總頁數 ({total_pages})！")
                return

            # 迭代所有頁面，只加入「不想」刪除的頁面
            for i in range(total_pages):
                if i not in pages_to_delete:
                    writer.add_page(reader.pages[i])

            # 寫入新檔案
            with open(output_path, "wb") as f:
                writer.write(f)

            messagebox.showinfo("成功", f"已成功刪除頁面！\n新檔案儲存於:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("處理失敗", f"讀取或寫入 PDF 時發生錯誤:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PdfDeleterApp(root)
    root.mainloop()