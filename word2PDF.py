import os
import win32com.client as win32
import tkinter as tk
from tkinter import messagebox

# 轉換主程式
def convert_word_to_pdf(input_folder, output_folder, text_widget, root):
    # 檢查輸入與輸出資料夾是否存在
    if not os.path.exists(input_folder):
        raise FileNotFoundError(f"輸入資料夾 {input_folder} 不存在")
    if output_folder == "":
        output_folder = os.path.join(input_folder, "pdf")
        os.makedirs(output_folder, exist_ok= True)
    elif not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 初始化 Word 應用程式
    word = win32.Dispatch('Word.Application')

    # 初始化計算器及計算處理檔案量
    processed_count = 0
    
    file_count = len([file for file in os.listdir(input_folder) 
                     if (file.endswith(".docx")
                         or file.endswith(".doc"))
                         and not file.startswith(".")
                         and not file.startswith("~")])
    
    #逐一處理資料夾內的 Word 文件
    for filename in os.listdir(input_folder):
        if ((filename.endswith(".docx") or 
            filename.endswith(".doc")) and not filename.startswith(".")
            and not filename.startswith("~")):
            doc_path = os.path.join(input_folder, filename)
            pdf_filename = os.path.splitext(filename)[0] + ".pdf"
            pdf_path = os.path.join(output_folder, pdf_filename)

            # 打開 Word 文件並將其另存為 PDF
            doc = word.Documents.Open(doc_path)
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 是 PDF 格式
            doc.Close()

            # 顯示進度
            processed_count += 1            
            progress_message = f"{filename} save as pdf file\n"
            text_widget.insert(tk.END, progress_message)
            #print(progress_message)
            progress_message = f"完成進度{round(processed_count/file_count*100)}%\n"
            root.update()

            text_widget.insert(tk.END, progress_message)
            #print(progress_message)
            text_widget.see(tk.END)
            root.update()
    
    #關閉 Word 應用程式
    word.Quit()


# 彈出式視窗，輸入來源資料等
def run_app():
    def onconvert():
        input_folder = input_folder_entry.get()
        output_folder = output_folder_entry.get()

        # 檢查來源資料夾是否為空
        if input_folder == "":
            messagebox.showerror("你沒輸入來源資料夾")
            return
        convert_buttom.config(state=tk.DISABLED)

        try:
            convert_word_to_pdf(input_folder, output_folder, progress_text, root)
        except Exception as e:
            messagebox.showerror("錯誤", str(e))

        convert_buttom.config(text="關閉程式，快下班", state=tk.NORMAL, command=root.quit)
    
    root = tk.Tk()
    root.title("wold 轉 pdf")

    # 框跟標籤
    tk.Label(root, text= "輸入資料夾路徑:").grid(row=0, column=0, padx=10, pady=5)
    input_folder_entry = tk.Entry(root, width=50)
    input_folder_entry.grid(row=0, column=1, padx=20, pady= 5)

    tk.Label(root, text= "輸出資料夾路徑(非必須):").grid(row=1, column=0, padx=10, pady=5)
    output_folder_entry = tk.Entry(root, width=50)
    output_folder_entry.grid(row=1, column=1, padx=20, pady= 5) 
    
    # text進度
    progress_text = tk.Text(root, height=10, width=60)
    progress_text.grid(row=3, column=1, padx=20, pady=10)

    # start按鈕
    convert_buttom = tk.Button(root, text="word開始轉換pdf", command= onconvert)
    convert_buttom.grid(row=2, column=1, pady=10)

    root.mainloop()

input_folder = ""
output_folder = ""
# input_folder = r"C:\Users\MAA\Desktop\新增資料夾\test"
# output_folder = ""

def __main__():
    run_app()


__main__()