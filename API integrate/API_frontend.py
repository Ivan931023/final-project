import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext
from PIL import Image, ImageTk
import requests

# API 伺服器 URL
API_BASE_URL = "http://127.0.0.1:5000"


class Operation:
    def __init__(self):
        self.stock_number = None
        self.start_date = None
        self.end_date = None
        self.data = None

    def update_fields(self):
        self.stock_number = stock_entry.get().strip()
        self.start_date = start_date_entry.get().strip()
        self.end_date = end_date_entry.get().strip()


    def display_chart(self, image_path):
        try:
            image = Image.open(image_path)
            chart_image = ImageTk.PhotoImage(image)
            chart_label.config(image=chart_image)
            chart_label.image = chart_image
        except Exception as e:
            messagebox.showerror("Error", f"無法顯示圖片：{e}")

    def submit_form(self):
        self.update_fields()
        if self.stock_number and self.start_date and self.end_date:
            try:
                response = requests.post(
                    f'{API_BASE_URL}/api/analyze_data',
                    json={'stock': self.stock_number, 'start': self.start_date, 'end': self.end_date}
                )
                response.raise_for_status()
                data = response.json()
                self.data = data
                
                if 'images' in data:
                    for i in range(4):
                        image_path = data['images'][i]
                        self.display_chart(image_path)
            except Exception as e:
                messagebox.showerror("Error", f"請求失敗：{e}")
        else:
            messagebox.showwarning("Input Error", "請填寫所有欄位！")
            
    def strategy_comparison(self):
        self.update_fields()
        if self.stock_number and self.start_date and self.end_date and self.data:
            try:
                if 'images' in self.data:
                    self.display_chart(self.data['images'][4])  
            except Exception as e:
                messagebox.showerror("Error", f"請求失敗：{e}")
        else:
            messagebox.showwarning("Input Error", "請輸入股票編號、起始時間和結束時間並查詢才能進行策略比較！")

    
    def gpt_response(self):
        self.update_fields()
        if self.stock_number and self.start_date and self.end_date and self.data:
            
            user_question = gpt_entry.get().strip()
            if not user_question:
                messagebox.showwarning("Input Error", "請輸入您的問題！")
                return
        
            try:
                response = requests.post(
                f'{API_BASE_URL}/api/send_to_gpt',
                json={'image_paths': [i for i in self.data['images']], "question": user_question}
                )
                response.raise_for_status()
                data = response.json()
                gpt_reply_text.delete(1.0, tk.END)  # 清空之前的內容
                gpt_reply_text.insert(tk.END, data.get('gpt_reply', 'AI 顧問無法提供回答'))
            except Exception as e:
                messagebox.showerror("Error", f"請求失敗：{e}")
        else:
            messagebox.showwarning("Input Error", "請輸入股票編號、起始時間和結束時間並查詢才能進行AI諮詢！")

    def download_excel(self):
        self.update_fields()
        if self.stock_number and self.start_date and self.end_date and self.data:
            try:
                excel_path = self.data.get('excel_path')
                if excel_path:
                    file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel Files", "*.xlsx")])
                    if file_path:
                        response = requests.get(f'{API_BASE_URL}/api/download_excel', params={'file_path': excel_path})
                        response.raise_for_status()
                        with open(file_path, 'wb') as file:
                            file.write(response.content)
                        messagebox.showinfo("Success", "分析報告已成功下載！")
            except Exception as e:
                messagebox.showerror("Error", f"請求失敗：{e}")
        else:
            messagebox.showwarning("Input Error", "請輸入股票編號、起始時間和結束時間並查詢才能下載分析報告！")

# GUI 介面設定
root = tk.Tk()
root.title("股票分析與AI諮詢應用程式")
root.geometry("600x600")

header_label = tk.Label(root, text="股票分析與AI諮詢應用程式", font=("標楷體", 20))
header_label.pack(pady=10)

form_frame = tk.Frame(root)
form_frame.pack(pady=10)

stock_label = tk.Label(form_frame, text="股票編號：")
stock_label.grid(row=0, column=0, padx=5)

stock_entry = tk.Entry(form_frame)
stock_entry.grid(row=0, column=1, padx=5)

start_date_label = tk.Label(form_frame, text="起始日期：")
start_date_label.grid(row=1, column=0, padx=5)

start_date_entry = tk.Entry(form_frame)
start_date_entry.grid(row=1, column=1, padx=5)

end_date_label = tk.Label(form_frame, text="結束日期：")
end_date_label.grid(row=2, column=0, padx=5)

end_date_entry = tk.Entry(form_frame)
end_date_entry.grid(row=2, column=1, padx=5)

op = Operation()

submit_button = tk.Button(form_frame, text="開始查詢", command=op.submit_form)
submit_button.grid(row=3, column=0, columnspan=3, pady=10)

chart_label = tk.Label(root)
chart_label.pack(pady=10)

strategy_button = tk.Button(root, text="進行策略比較", command=op.strategy_comparison)
strategy_button.pack(pady=10)

gpt_label = tk.Label(root, text="請輸入您的問題：")
gpt_label.pack(pady=5)

gpt_entry = tk.Entry(root, width=50)
gpt_entry.pack(pady=5)

gpt_button = tk.Button(root, text="發送問題至AI顧問", command=op.gpt_response)
gpt_button.pack(pady=5)

gpt_reply_label = tk.Label(root, text="（AI顧問回應將顯示在此處）")
gpt_reply_label.pack(pady=10)
gpt_reply_label.destroy()  # 刪除原有的 gpt_reply_label
gpt_reply_text = scrolledtext.ScrolledText(root, width=70, height=10, wrap=tk.WORD)
gpt_reply_text.pack(pady=10)

excel_button = tk.Button(root, text="下載分析結果", command=op.download_excel)
excel_button.pack(pady=10)

root.mainloop()