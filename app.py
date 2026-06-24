import os
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
import pandas as pd

# CustomTkinterの基本設定
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.TkinterDnD.DnDWrapper.__init__(self)

        self.title("CSVデータ抽出ツール")
        self.geometry("500x350")

        self.setup_ui()

    def setup_ui(self):
        # メインフレーム
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(padx=20, pady=20, fill="both", expand=True)

        # タイトルラベル
        self.title_label = ctk.CTkLabel(
            self.frame, 
            text="CSVファイルをここにドラッグ＆ドロップしてください", 
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.title_label.pack(padx=10, pady=20)

        # ドラッグ＆ドロップの受け入れ領域
        self.drop_area = ctk.CTkLabel(
            self.frame,
            text="\n\n ここにファイルをドロップ \n\n",
            fg_color=("#E0E0E0", "#333333"),
            corner_radius=10,
            width=400,
            height=120
        )
        self.drop_area.pack(padx=10, pady=10)

        # D&Dの登録
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.handle_drop)

        # ボタンで選択
        self.select_button = ctk.CTkButton(
            self.frame, 
            text="ファイルを手動で選択", 
            command=self.select_file
        )
        self.select_button.pack(padx=10, pady=15)

        # ステータス表示
        self.status_label = ctk.CTkLabel(self.frame, text="", text_color="green")
        self.status_label.pack(padx=10, pady=5)

    def select_file(self):
        """手動でファイルを選択したときの処理"""
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.process_csv(file_path)

    def handle_drop(self, event):
        """ドラッグ＆ドロップされたときの処理"""
        file_path = event.data
        
        # WindowsのD&Dでパスに括弧 {} がつく場合があるのを外す
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
            
        if file_path.lower().endswith('.csv'):
            self.process_csv(file_path)
        else:
            messagebox.showerror("エラー", "CSVファイルのみ受け付けています。")

    def process_csv(self, file_path):
        """CSVの抽出処理とユーザー指定先への保存"""
        try:
            # 1. CSVの読み込み（Shift-JIS標準）
            df = pd.read_csv(file_path, encoding='shift_jis')

            # --- 🛠️ ここに独自のデータ抽出・加工ロジックを記述します ---
            # 例: processed_df = df[df['列名'] == '条件']
            processed_df = df.copy() 
            # --------------------------------------------------

            # 2. 初期ファイル名の設定（「抽出_元のファイル名.csv」を提案）
            base_name = os.path.basename(file_path)
            default_output_name = f"抽出_{base_name}"

            # 3. 保存先を選択するダイアログを表示
            output_path = filedialog.asksaveasfilename(
                initialfile=default_output_name,
                defaultextension=".csv",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
                title="処理後のファイルの保存先を選択してください"
            )

            # ユーザーがキャンセルした場合は処理を中断
            if not output_path:
                self.status_label.configure(text="保存がキャンセルされました", text_color="orange")
                return

            # 4. 指定された場所に保存
            processed_df.to_csv(output_path, index=False, encoding='shift_jis')
            
            # 画面上の表示を更新
            saved_filename = os.path.basename(output_path)
            self.status_label.configure(text=f"保存完了: {saved_filename}", text_color="green")
            messagebox.showinfo("成功", f"ファイルを保存しました：\n{saved_filename}")

        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの処理中にエラーが発生しました:\n{str(e)}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
