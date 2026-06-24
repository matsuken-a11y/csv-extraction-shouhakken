import os
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES
import pandas as pd

# CustomTkinterの基本設定
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# エラー回避のため、ベースにTkinterDnD.Tkの仕組みを持つ特殊なクラスを作ります
class CustomCTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # ドラッグ＆ドロップの仕組みを確実に初期化
        self.TkinterDnD.DnDWrapper.__init__(self)

class App:
    def __init__(self):
        # メインウィンドウの作成
        self.root = CustomCTk()
        self.root.title("CSVデータ抽出ツール")
        self.root.geometry("500x350")

        self.setup_ui()

    def setup_ui(self):
        # メインフレーム
        self.frame = ctk.CTkFrame(self.root)
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
        self.status_label = ctk.CTkLabel(self.root, text="", text_color="green")
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

            # --- 🛠️ 【VBAの移植部分】データ抽出・加工ロジック ---
            # ⚠️注意：CSVの1行目にある実際の「列の見出し名」と一致させてください。
            # もしCSVの見出しが F, AF, AH 等ではなく日本語見出しの場合は、その文字列に書き換えてください。
            
            # 【条件1の抽出】 F列=1、AF列=102003、AH列=21102010
            cond1 = (df['F'] == 1) & (df['AF'] == 102003) & (df['AH'] == 21102010)
            df_cond1 = df[cond1][['I', 'N', 'O', 'AY', 'AW', 'AX']].reset_index(drop=True)
            df_cond1.columns = ['A', 'B', 'C', 'D', 'E', 'F'] # 左側の列(A~F)に割り当て

            # 【条件2の抽出】 F列=2、AF列=102003、AH列=21102010
            cond2 = (df['F'] == 2) & (df['AF'] == 102003) & (df['AH'] == 21102010)
            df_cond2 = df[cond2][['I', 'N', 'O', 'AY', 'AW', 'AX']].reset_index(drop=True)
            df_cond2.columns = ['J', 'K', 'L', 'M', 'N_out', 'O_out'] # 右側の列(J~O)に割り当て

            # --- 💡 【Excel数式の自動計算部分】G列・H列・I列の作成 ---
            
            # 空のG, H, I列を左側のデータと同じ行数分、初期化
            df_headers = pd.DataFrame(index=df_cond1.index)
            df_headers['G'] = ""
            df_headers['H'] = ""
            df_headers['I'] = ""

            # 右側の O_out (ExcelでのO列) にある値をセットとして取得（高速検索用）
            o_set = set(df_cond2['O_out'].dropna())

            # 右側の O_out と K (ExcelでのK列) のマッピング辞書を作成（H列のINDEX+MATCH用）
            mapping_k = dict(zip(df_cond2['O_out'], df_cond2['K']))
            
            # 右側の O_out と L (ExcelでのL列) のマッピング辞書を作成（I列のINDEX+MATCH用）
            mapping_l = dict(zip(df_cond2['O_out'], df_cond2['L']))

            # 1行ずつG列、H列、I列の値を計算
            for idx, row in df_cond1.iterrows():
                f_val = row['F']
                
                # G3の数式に相当：F列の値が右側のO列にあれば「〇」
                if f_val in o_set:
                    df_headers.at[idx, 'G'] = "〇"
                    
                    # H3の数式に相当：一致した右側O列の行にあるK列の値を引っ張る
                    df_headers.at[idx, 'H'] = mapping_k.get(f_val, "")
                    
                    # I3の数式に相当：一致した右側O列の行にあるL列の値を引っ張る
                    df_headers.at[idx, 'I'] = mapping_l.get(f_val, "")
                else:
                    df_headers.at[idx, 'G'] = ""
                    df_headers.at[idx, 'H'] = ""
                    df_headers.at[idx, 'I'] = ""

            # 【すべての列を結合】 A~F, G~I, J~O を横に綺麗に並べる
            processed_df = pd.concat([df_cond1, df_headers, df_cond2], axis=1)
            
            # CSV出力時に見やすいよう、列の順番を固定
            column_order = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N_out', 'O_out']
            processed_df = processed_df.reindex(columns=column_order)
            
            # 最後に、Excelで見出しが分かりやすいよう通常のアルファベット名に整える
            processed_df.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']
            # --------------------------------------------------

            # 2. 初期ファイル名の設定
            base_name = os.path.basename(file_path)
            default_output_name = f"松柏軒_{base_name}"

            # 3. 保存先を選択するダイアログを表示
            output_path = filedialog.asksaveasfilename(
                initialfile=default_output_name,
                defaultextension=".csv",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
                title="処理後のファイルの保存先を選択してください"
            )

            if not output_path:
                self.status_label.configure(text="保存がキャンセルされました", text_color="orange")
                return

            # 4. 指定された場所に保存（Excelで開けるようにShift-JIS、インデックス無し）
            processed_df.to_csv(output_path, index=False, encoding='shift_jis')
            
            saved_filename = os.path.basename(output_path)
            self.status_label.configure(text=f"保存完了: {saved_filename}", text_color="green")
            messagebox.showinfo("成功", f"ファイルを保存しました：\n{saved_filename}")

        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの処理中にエラーが発生しました:\n{str(e)}")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = App()
    app.run()
