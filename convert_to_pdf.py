import sys
import os
import win32com.client

def convert_excel_to_pdf(input_path):
    """ExcelのアクティブシートのみをPDFに変換"""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(input_path)
        ws = wb.ActiveSheet  # アクティブなシートのみ
        pdf_path = os.path.splitext(input_path)[0] + ".pdf"
        
        ws.ExportAsFixedFormat(0, pdf_path, IgnorePrintAreas=False)  # アクティブシートのみ
        wb.Close(False)
        return pdf_path
    except Exception as e:
        print(f"Excel変換エラー: {e}")
    finally:
        excel.Quit()

def convert_word_to_pdf(input_path):
    """WordファイルをPDFに変換"""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(input_path)
        pdf_path = os.path.splitext(input_path)[0] + ".pdf"
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close(False)
        return pdf_path
    except Exception as e:
        print(f"Word変換エラー: {e}")
    finally:
        word.Quit()

def main():
    if len(sys.argv) < 2:
        print("ファイルをドラッグ＆ドロップしてください")
        return

    for file_path in sys.argv[1:]:
        if not os.path.exists(file_path):
            print(f"ファイルが見つかりません: {file_path}")
            continue

        ext = os.path.splitext(file_path)[1].lower()
        if ext in [".xlsx", ".xlsm"]:
            pdf_path = convert_excel_to_pdf(file_path)
            print(f"Excel PDF作成: {pdf_path}")
        elif ext in [".docx", ".doc"]:
            pdf_path = convert_word_to_pdf(file_path)
            print(f"Word PDF作成: {pdf_path}")
        else:
            print(f"対応していないファイル形式: {ext}")

if __name__ == "__main__":
    main()
