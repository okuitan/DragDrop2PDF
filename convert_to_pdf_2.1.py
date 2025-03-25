import sys
import os
import win32com.client
import tkinter as tk
from tkinter import simpledialog

def convert_excel_to_pdf(input_path, orientation, paper_size):
    """ExcelのアクティブシートのみをPDFに変換"""
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        # Excelを非表示に設定（バッチ処理用）
        excel.Visible = False

        wb = excel.Workbooks.Open(input_path)
        ws = wb.ActiveSheet  # アクティブなシートのみ
        
        # ページ設定で印刷方向を設定
        ws.PageSetup.Orientation = orientation  # 1: 縦（ポートレート）、2: 横（ランドスケープ）
        
        # 用紙サイズの設定
        if paper_size == 1:  # A4
            ws.PageSetup.PaperSize = 9  # A4の定数
        elif paper_size == 2:  # A3
            ws.PageSetup.PaperSize = 8  # A3の定数
        elif paper_size == 3:  # A5
            ws.PageSetup.PaperSize = 7  # A5の定数
        
        # 拡大縮小設定（用紙にぴったり合わせる）
        ws.PageSetup.FitToPagesWide = 1  # 1ページの幅に合わせる
        ws.PageSetup.FitToPagesTall = 1  # 1ページの高さに合わせる

        # 保存先のファイル名を決定
        if orientation == 1:
            pdf_path = os.path.splitext(input_path)[0] + "_portrait.pdf"
        else:
            pdf_path = os.path.splitext(input_path)[0] + "_landscape.pdf"
        
        ws.ExportAsFixedFormat(0, pdf_path, IgnorePrintAreas=False)  # アクティブシートのみ
        wb.Close(False)
        return pdf_path
    except Exception as e:
        print(f"Excel変換エラー: {e}")
    finally:
        # excelが正常にインスタンス化されていればQuit()を呼び出す
        if excel is not None:
            try:
                excel.Quit()  # Excelを終了
            except Exception as e:
                print(f"Excel終了エラー: {e}")

def ask_orientation():
    """ダイアログで縦か横を選ばせる"""
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを表示しない
    result = simpledialog.askstring("選択", "縦（Portrait）か横（Landscape）を入力してください（1: 縦、2: 横）")
    
    # 入力が1または2でなければ、デフォルトで縦（1）を設定
    if result == "2":
        return 2  # 横（ランドスケープ）
    return 1  # 縦（ポートレート）

def ask_paper_size():
    """ダイアログで用紙サイズを選ばせる"""
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを表示しない
    result = simpledialog.askstring("選択", "用紙サイズを選んでください（1: A4、2: A3、3: A5）")
    
    if result == "2":
        return 2  # A3
    elif result == "3":
        return 3  # A5
    return 1  # A4

def main():
    if len(sys.argv) < 2:
        print("ファイルをドラッグ＆ドロップしてください")
        return

    # 印刷方向をダイアログで聞く
    orientation = ask_orientation()

    # 用紙サイズをダイアログで聞く
    paper_size = ask_paper_size()

    for file_path in sys.argv[1:]:
        if not os.path.exists(file_path):
            print(f"ファイルが見つかりません: {file_path}")
            continue

        ext = os.path.splitext(file_path)[1].lower()
        if ext in [".xlsx", ".xlsm"]:
            pdf_path = convert_excel_to_pdf(file_path, orientation, paper_size)
            if orientation == 1:
                print(f"Excel PDF作成（縦向き、用紙サイズ: {paper_size}）: {pdf_path}")
            else:
                print(f"Excel PDF作成（横向き、用紙サイズ: {paper_size}）: {pdf_path}")
        else:
            print(f"対応していないファイル形式: {ext}")

if __name__ == "__main__":
    main()
