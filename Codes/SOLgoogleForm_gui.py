import tkinter as tk
from tkinter import colorchooser,filedialog,font,messagebox
import openpyxl as op
import SOLgoogleForm_excel as sgf

class googleFormGui:
    def __init__(self,master):
        self.master=master
        self.master.title("Read Google Form")
        self.master.geometry("600x100")

        #excel生成用に渡す引数
        self.open_dir=None
        self.save_dir=None
        #表示する色の初期化


        ###event処理達
        def generate_event():
            #excelシートの生成
            wb=op.Workbook()
            wb.save(self.save_dir)
            m=sgf.sfg(self.open_dir,self.save_dir)
            #messageBox
            self.mb=messagebox.showinfo("通知","生成完了")

        def open_select_file():
            file_type = [("読込Excel", "*.xlsx")]
            selected_file = filedialog.askopenfilename(filetypes=file_type,defaultextension="xlsx")
            if selected_file:
                # for file in selected_files:
                #     self.listbox_files.insert(tk.END, file)
                self.open_dir=selected_file
                self.open_dir_label.config(text=selected_file)
        
        def save_select_file():
            file_type = [("保存Excel", "*.xlsx")]
            selected_file = filedialog.asksaveasfilename(filetypes=file_type,defaultextension="xlsx")
            if selected_file:
                # for file in selected_files:
                #     self.listbox_files.insert(tk.END, file)
                self.save_dir=selected_file
                self.save_dir_label.config(text=selected_file)
                

        # #gridの行列はこれで取得できる(デバッグ用)
        # def callback(event):
        #     info = event.widget.grid_info()
        #     print(info['row'],info['column'])

        #widget
        
        #label
        self.open_label=tk.Label(text="スプレッドシート")
        self.save_label=tk.Label(text="保存先")
        self.open_dir_label=tk.Label(text="読み込むファイルを選択(.xlsx)")
        self.save_dir_label=tk.Label(text="保存ファイルを選択(.xlsx)")
        #entry

        #button
        self.open_selectFile_button=tk.Button(text="参照",command=open_select_file)
        self.save_selectFile_button=tk.Button(text="参照",command=save_select_file)
        self.generate_button=tk.Button(text="generate",command=generate_event)
        #scroll,listbox

        #gridで配置
        self.open_label.grid(row=0,column=0,sticky="e")
        self.save_label.grid(row=1,column=0,sticky="e")
        self.open_selectFile_button.grid(row=0,column=1,sticky="e")
        self.save_selectFile_button.grid(row=1,column=1,sticky="e")
        self.open_dir_label.grid(row=0,column=2,sticky="w")
        self.save_dir_label.grid(row=1,column=2,sticky="w")
        self.generate_button.grid(row=3,column=0,sticky="w")

        self.master.grid_columnconfigure(0,weight=1,minsize=100)    
        self.master.grid_columnconfigure(1,weight=1,minsize=100)    
        self.master.grid_columnconfigure(2,weight=1,minsize=100)    
        self.master.grid_columnconfigure(3,weight=1,minsize=100)    
        self.master.grid_rowconfigure(7,weight=1,minsize=100)

        # self.master.bind_all("<1>",callback) #gridの行列取得


if __name__ == '__main__':
    root = tk.Tk()
    make = googleFormGui(root)
    root.mainloop()       
        