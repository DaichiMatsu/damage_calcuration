import tkinter 
from tkinter import ttk
from tkinter import filedialog
import openpyxl
from pathlib import Path
import math

class Application(tkinter.Frame):
    def __init__(self, root=None):
        super().__init__(root,
                         width=420, height=720,
                         borderwidth=1, relief="groove")
        self.root = root
        self.pack() # 位置を設定して配置
        self.pack_propagate(0) # サイズ調整
        self.create_widgets()


    def create_widgets(self):
        # 閉じるボタン
        quit_btn = tkinter.Button(self)
        quit_btn["text"] = "閉じる"
        quit_btn["command"] = self.root.destroy
        quit_btn.pack(side="bottom") # 位置指定　bottom=下

        self.cell_positions = []
        self.cell_texts = []

        self.message = tkinter.Message(self)
        self.message.place(x="180", y=310)

        self.loading_message = tkinter.Message(self)
        self.loading_message.place(x="180", y=310)

    
        # メッセージ出力(A/C)
        text_label = tkinter.Message(self)
        text_label["text"] = "A/C"
        text_label.place(x=10, y=str(22))
        # テキストボックス(A/C)
        self.text_box_A_C = tkinter.Entry(self)
        self.text_box_A_C["width"] = 10
        self.text_box_A_C.place(x="50", y=str(22))
        self.cell_texts.append(self.text_box_A_C)   

        # メッセージ出力(努力値A/C)
        text_label = tkinter.Message(self)
        text_label["text"] = "A/C努力値"
        text_label.place(x=10, y=str(2*22))
        # テキストボックス(努力値A/C)
        self.text_box_efA_C = tkinter.Entry(self)
        self.text_box_efA_C["width"] = 10
        self.text_box_efA_C.place(x="50", y=str(2*22))
        self.cell_texts.append(self.text_box_efA_C) 

        # メッセージ出力(わざ)
        text_label = tkinter.Message(self)
        text_label["text"] = "わざ"
        text_label.place(x=10, y=str(3*22))
        # テキストボックス(わざ)
        self.text_box_mo = tkinter.Entry(self)
        self.text_box_mo["width"] = 10
        self.text_box_mo.place(x="50", y=str(3*22))
        self.cell_texts.append(self.text_box_mo)   

        # メッセージ出力(タイプ一致)
        text_label = tkinter.Message(self)
        text_label["text"] = "タイプ一致"
        text_label.place(x=10, y=str(4*22))
        # チェックボックス(タイプ一致)
        self.checked_type_same = tkinter.BooleanVar()
        self.check_box_type_same = tkinter.Checkbutton(
            self,
            variable=self.checked_type_same
        )
        self.check_box_type_same.place(x=100, y=100)
        self.check_box_type_same.place(x="50", y=str(4*22))


        # メッセージ出力(B/D)
        text_label = tkinter.Message(self)
        text_label["text"] = "B/D"
        text_label.place(x=200, y=str(22))
        # テキストボックス(B/D)
        self.text_box_B_D = tkinter.Entry(self)
        self.text_box_B_D["width"] = 10
        self.text_box_B_D.place(x="240", y=str(22))
        self.cell_texts.append(self.text_box_B_D)   

        # メッセージ出力(HP)
        text_label = tkinter.Message(self)
        text_label["text"] = "HP"
        text_label.place(x=200, y=str(2*22))
        # テキストボックス(HP)
        self.text_box_HP = tkinter.Entry(self)
        self.text_box_HP["width"] = 10
        self.text_box_HP.place(x="240", y=str(2*22))
        self.cell_texts.append(self.text_box_HP) 

        # メッセージ出力(努力値B/D)
        text_label = tkinter.Message(self)
        text_label["text"] = "B/D努力値"
        text_label.place(x=200, y=str(3*22))
        # テキストボックス(努力値B/D)
        self.text_box_efB_D = tkinter.Entry(self)
        self.text_box_efB_D["width"] = 10
        self.text_box_efB_D.place(x="240", y=str(3*22))
        self.cell_texts.append(self.text_box_efB_D) 

        # メッセージ出力(努力値HP)
        text_label = tkinter.Message(self)
        text_label["text"] = "HP努力値"
        text_label.place(x=200, y=str(4*22))
        # テキストボックス(努力値HP)
        self.text_box_efHP = tkinter.Entry(self)
        self.text_box_efHP["width"] = 10
        self.text_box_efHP.place(x="240", y=str(4*22))
        self.cell_texts.append(self.text_box_efHP) 

        # メッセージ出力(タイプ相性)
        text_label = tkinter.Message(self)
        text_label["text"] = "タイプ相性"
        text_label.place(x=200, y=str(5*22))
        # テキストボックス(タイプ相性)
        self.text_box_type_match = ttk.Combobox(
        self,
        values=["等倍", "ちょうばつぐん", "ばつぐん", "いまひとつ", "かなりいまひとつ"],  # 選択肢
        state="readonly",           # 入力不可（選択のみ）
        width=10
        )
        self.text_box_type_match.current(0)  # 最初の項目を選択状態にする
        self.text_box_type_match.place(x="240", y=str(5*22))


        # 実行ボタン
        submit_btn = tkinter.Button(self)
        submit_btn["text"] = "保存"
        submit_btn["command"] = self.save_data
        submit_btn.place(x=150, y=230)

        # 計算ボタン
        calc_btn = tkinter.Button(self)
        calc_btn["text"] = "計算"
        calc_btn["command"] = self.calculation
        calc_btn.place(x=150, y=600)

        # リセットボタン
        reset_btn = tkinter.Button(self)
        reset_btn["text"] = "リセット"
        reset_btn["command"] = self.reset
        reset_btn.place(x=150, y=400)
        

        # 読み込みボタン
        submit_btn = tkinter.Button(self)
        submit_btn["text"] = "読み込み"
        submit_btn["command"] = self.loading_data
        submit_btn.place(x=150, y=270)

    # 実行ボタンを押すとテキストを入力するシステムをメソッドとして作成
    def save_data(self):
        wb = openpyxl.load_workbook("app_data.xlsx") # 選択したファイルを読み込み
        ws = wb.worksheets[0] # ワークシートの一枚目(0番目)をwsに保存
        text_A_C = self.text_box_A_C.get()
        ws["A1"].value = text_A_C
        text_efA_C = self.text_box_efA_C.get()
        ws["A2"].value = text_efA_C
        text_mo = self.text_box_mo.get()
        ws["A3"].value = text_mo
        check_box_type_same = self.checked_type_same.get()
        ws["A4"].value = check_box_type_same


        text_B_D = self.text_box_B_D.get()
        ws["B1"].value = text_B_D
        text_HP = self.text_box_HP.get()
        ws["B2"].value = text_HP
        text_efB_D = self.text_box_efB_D.get()
        ws["B3"].value = text_efB_D
        text_efHP = self.text_box_efHP.get()
        ws["B4"].value = text_efHP
        text_box_type_match = self.text_box_type_match.get()
        ws["B5"].value = text_box_type_match
        
        
        wb.save("app_data.xlsx") # book保存
        self.message["text"] = "保存完了"

    def calculation(self):
        wb = openpyxl.load_workbook("app_data.xlsx")
        ws = wb.worksheets[0]

        # 文字列をint型に直す
        A_C = int(ws["A1"].value)
        efA_C = int(ws["A2"].value)
        mo = int(ws["A3"].value)
        B_D = int(ws["B1"].value)
        # HP = int(ws["B2"].value)
        efB_D = int(ws["B3"].value)
        # efHP = int(ws["B4"].value)

        # 計算結果保存
        rand_min = 0.85
        base_damage = 22 * mo * (A_C+efA_C) / (B_D+efB_D)

        base_damage_tr = math.floor(base_damage)

        damage = (base_damage_tr / 50) +2

        # 最大ダメージ
        damage_tr = math.floor(damage)

        damage_rand_min = damage_tr * rand_min

        # 最小ダメージ
        damage_rand_min_tr = math.floor(damage_rand_min)

        ws["Z1"].value = damage_tr
        wb.save("app_data.xlsx")

        self.loading_message["text"] = damage_tr, damage_rand_min_tr

    def reset(self):
        for i in self.cell_texts:
            i.delete(0, tkinter.END)
        self.checked_type_same.set(0)
        self.text_box_type_match.set(0)
        
        wb = openpyxl.load_workbook("app_data.xlsx") # 選択したファイルを読み込み
        ws = wb.worksheets[0] # ワークシートの一枚目(0番目)をwsに保存
        text_A_C = self.text_box_A_C.get()
        ws["A1"].value = text_A_C
        text_efA_C = self.text_box_efA_C.get()
        ws["A2"].value = text_efA_C
        text_mo = self.text_box_mo.get()
        ws["A3"].value = text_mo
        check_box_type_same = self.checked_type_same.get()
        ws["A4"].value = check_box_type_same


        text_B_D = self.text_box_B_D.get()
        ws["B1"].value = text_B_D
        text_HP = self.text_box_HP.get()
        ws["B2"].value = text_HP
        text_efB_D = self.text_box_efB_D.get()
        ws["B3"].value = text_efB_D
        text_efHP = self.text_box_efHP.get()
        ws["B4"].value = text_efHP
        text_box_type_match = self.text_box_type_match.get()
        ws["B5"].value = text_box_type_match
        
        
        wb.save("app_data.xlsx") # book保存
        self.message["text"] = "リセット"
    
    # 読み込みボタンを押すとエクセル内容が表示されるメソッドを作成
    def loading_data(self):
        wb = openpyxl.load_workbook("app_data.xlsx")
        ws = wb.worksheets[0] 
        values = [] # cellの中の文字を入れるためのリスト
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None: # cellの中に文字がある時にリストにその文字を追加する
                    values.append(cell.value)
        self.loading_message["text"] = values

root = tkinter.Tk()
root.title("さぷー アプリ")
root.geometry("450x750")
app = Application(root=root)
root.mainloop()