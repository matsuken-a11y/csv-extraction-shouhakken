# csv-extraction-shouhakkenimport os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd


class CSVConverterApp:

    def __init__(self, root):
        self.root = root
        self.root.title("CSV条件抽出ツール")
        self.root.geometry("500x250")
        self.root.resizable(False, False)

        # スタイル設定
        style = ttk.Style()
        style.theme_use("vista")

        # メインフレーム
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # タイトルラベル
        title_label = ttk.Label(
            main_frame,
            text="CSV原本読込 ＆ 松柏軒 転記ツール",
            font=("Helvetica", 14, "bold"),
        )
        title_label.pack(pady=10)

        # 説明文
        desc_label = ttk.Label(
            main_frame,
            text="「実行する」ボタンを押し、対象のCSVファイルを選択してください。\n※出力先のExcelファイルと同じフォルダに配置して実行してください。",
            justify=tk.CENTER,
        )
        desc_label.pack(pady=10)

        # 実行ボタン
        self.btn_run = ttk.Button(
            main_frame, text=" 実行する ", command=self.process_data, width=20
        )
        self.btn_run.pack(pady=15)

        # ステータスバー
        self.status_var = tk.StringVar(value="準備完了")
        status_bar = ttk.Label(
            root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding=(5, 2),
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def process_data(self):
        # ① CSVファイル選択ダイアログ
        csv_path = filedialog.askopenfilename(
            title="CSVファイルを選択してください",
            filetypes=[("CSVファイル", "*.csv")],
        )
        if not csv_path:
            messagebox.showinfo("情報", "処理をキャンセルしました。")
            return

        self.status_var.set("処理中...")
        self.btn_run.config(state=tk.DISABLED)
        self.root.update()

        # 出力先Excelファイル（アプリと同じフォルダにある「データ集計ブック.xlsx」を想定）
        # EXE化した場合、実行ファイルがあるフォルダを起点にします
        if getattr(sys, "frozen", False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        excel_path = os.path.join(base_dir, "データ集計ブック.xlsx")

        if not os.path.exists(excel_path):
            messagebox.showerror(
                "エラー",
                f"出力先のExcelファイルが見つかりません。\nアプリと同じフォルダに「データ集計ブック.xlsx」を配置してください。\n\n確認先: {excel_path}",
            )
            self.status_var.set("エラー：ファイル未検出")
            self.btn_run.config(state=tk.NORMAL)
            return

        try:
            # ② CSVの読み込み（Shift_JIS/CP932を想定）
            df_src = pd.read_csv(csv_path, encoding="cp932")

            # 列インデックスの動的取得
            col_F = df_src.columns[5]
            col_I = df_src.columns[8]
            col_N = df_src.columns[13]
            col_O = df_src.columns[14]
            col_AF = df_src.columns[31]
            col_AH = df_src.columns[33]
            col_AW = df_src.columns[48]
            col_AX = df_src.columns[49]
            col_AY = df_src.columns[50]

            # 型を数値に変換（VBAのVal関数に相当）
            df_src[col_F] = pd.to_numeric(df_src[col_F], errors="coerce").fillna(
                0
            )
            df_src[col_AF] = pd.to_numeric(
                df_src[col_AF], errors="coerce"
            ).fillna(0)
            df_src[col_AH] = pd.to_numeric(
                df_src[col_AH], errors="coerce"
            ).fillna(0)

            # ③ 条件抽出処理
            # --- 条件1の抽出 (F=1, AF=102003, AH=21102010)
            cond1 = (
                (df_src[col_F] == 1)
                & (df_src[col_AF] == 102003)
                & (df_src[col_AH] == 21102010)
            )
            df_dst1 = df_src.loc[
                cond1, [col_I, col_N, col_O, col_AY, col_AW, col_AX]
            ]

            # --- 条件2の抽出 (F=2, AF=102003, AH=21102010)
            cond2 = (
                (df_src[col_F] == 2)
                & (df_src[col_AF] == 102003)
                & (df_src[col_AH] == 21102010)
            )
            df_dst2 = df_src.loc[
                cond2, [col_I, col_N, col_O, col_AY, col_AW, col_AX]
            ]

            # ④ Excelファイルへの書き込み
            with pd.ExcelWriter(
                excel_path,
                engine="openpyxl",
                mode="a",
                if_sheet_exists="overlay",
            ) as writer:
                # 【原本】シートへの全上書き
                df_src.to_excel(
                    writer,
                    sheet_name="原本",
                    index=False,
                    header=True,
                    startrow=0,
                )

                # 【松柏軒】シートへの上書き（A3セルから）
                df_dst1.to_excel(
                    writer,
                    sheet_name="松柏軒",
                    index=False,
                    header=False,
                    startrow=2,
                    startcol=0,
                )

                # 【松柏軒】シートへの上書き（J3セルから）
                df_dst2.to_excel(
                    writer,
                    sheet_name="松柏軒",
                    index=False,
                    header=False,
                    startrow=2,
                    startcol=9,
                )

            self.status_var.set("処理完了")
            messagebox.showinfo(
                "完了", "CSVの読込および条件抽出転記が完了しました。"
            )

        except Exception as e:
            self.status_var.set("エラー発生")
            messagebox.showerror(
                "エラー", f"エラーが発生しました。\n内容: {str(e)}"
            )

        finally:
            self.btn_run.config(state=tk.NORMAL)


if __name__ == "__main__":
    root = tk.Tk()
    app = CSVConverterApp(root)
    root.mainloop()
