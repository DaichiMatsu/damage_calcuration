# テキストボックス(物理/特殊)
        self.text_box_mo_type = ttk.Combobox(
        self,
        values=["物理", "特殊"],  # 選択肢
        state="readonly",           # 入力不可（選択のみ）
        width=10
        )
        self.text_box_mo_type.current(0)  # 最初の項目を選択状態にする
        self.text_box_mo_type.place(x="640", y=str(2*22))