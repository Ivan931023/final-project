import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk
import requests

# API 伺服器 URL
API_BASE_URL = "http://127.0.0.1:5000"

# 顯示圖表函數
def display_chart(image_path):
    try:
        image = Image.open(image_path)
        chart_image = ImageTk.PhotoImage(image)
        chart_label.config(image=chart_image)
        chart_label.image = chart_image
    except Exception as e:
        messagebox.showerror("Error", f"無法顯示圖片：{e}")

# 提交表單函數
def submit_form():
    stock_number = stock_entry.get().strip()
    start_date = start_date_entry.get().strip()
    end_date = end_date_entry.get().strip()
    
    if stock_number and start_date and end_date:
        try:
            response = requests.post(
                f'{API_BASE_URL}/api/analyze_data',
                json={'stock': stock_number, 'start': start_date, 'end': end_date}
            )
            response.raise_for_status()
            data = response.json()

            if 'images' in data:
                for i in range(4):
                    image_path = data['images'][i]
                    display_chart(image_path)
        except Exception as e:
            messagebox.showerror("Error", f"請求失敗：{e}")
    else:
        messagebox.showwarning("Input Error", "請填寫所有欄位！")

# 策略比較函數
def strategy_comparison():
    stock_number = stock_entry.get().strip()
    start_date = start_date_entry.get().strip()
    end_date = end_date_entry.get().strip()
    
    if stock_number and start_date and end_date:
        try:
            response = requests.post(
                f'{API_BASE_URL}/api/analyze_data',
                json={'stock': stock_number, 'start': start_date, 'end': end_date}
            )
            response.raise_for_status()
            data = response.json()

            if 'images' in data:
                display_chart(data['images'][4])  # 策略比較圖
        except Exception as e:
            messagebox.showerror("Error", f"請求失敗：{e}")
    else:
        messagebox.showwarning("Input Error", "請輸入股票編號、起始時間和結束時間才能進行策略比較！")

# GPT 交互函數
def gpt_response():
    stock_number = stock_entry.get().strip()
    start_date = start_date_entry.get().strip()
    end_date = end_date_entry.get().strip()
    
    if stock_number and start_date and end_date:
        user_question = gpt_entry.get().strip()
        if not user_question:
            messagebox.showwarning("Input Error", "請輸入您的問題！")
            return
    
        try:
            response = requests.post(
                f'{API_BASE_URL}/api/send_to_gpt',
                json={'image_paths': ['images/{}_Close.png'.format(stock_number)], "question": user_question}
            )
            response.raise_for_status()
            data = response.json()
            gpt_reply_label.config(text=data.get('gpt_reply', 'AI 顧問無法提供回答'))
        except Exception as e:
            messagebox.showerror("Error", f"請求失敗：{e}")
    else:
        messagebox.showwarning("Input Error", "請輸入股票編號、起始時間和結束時間才能進行AI諮詢！")

# 下載分析報告函數
def download_excel():
    stock_number = stock_entry.get().strip()
    start_date = start_date_entry.get().strip()
    end_date = end_date_entry.get().strip()
    
    if stock_number and start_date and end_date:
        try:
            response = requests.post(
                f'{API_BASE_URL}/api/analyze_data',
                json={'stock': stock_number, 'start': start_date, 'end': end_date}
            )
            response.raise_for_status()
            data = response.json()

            excel_path = data.get('excel_path')
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
        messagebox.showwarning("Input Error", "請輸入股票編號才能下載分析報告！")

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

submit_button = tk.Button(form_frame, text="開始查詢", command=submit_form)
submit_button.grid(row=3, column=0, columnspan=3, pady=10)

chart_label = tk.Label(root)
chart_label.pack(pady=10)

strategy_button = tk.Button(root, text="進行策略比較", command=strategy_comparison)
strategy_button.pack(pady=10)

gpt_label = tk.Label(root, text="請輸入您的問題：")
gpt_label.pack(pady=5)

gpt_entry = tk.Entry(root, width=50)
gpt_entry.pack(pady=5)

gpt_button = tk.Button(root, text="發送問題至AI顧問", command=gpt_response)
gpt_button.pack(pady=5)

gpt_reply_label = tk.Label(root, text="(AI顧問回應將顯示在此處)")
gpt_reply_label.pack(pady=10)

excel_button = tk.Button(root, text="下載分析結果", command=download_excel)
excel_button.pack(pady=10)

root.mainloop()