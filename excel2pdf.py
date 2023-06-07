import os
from tqdm import tqdm
import argparse
from win32com.client import DispatchEx
import atexit
from PyPDF2 import PdfMerger


def mkdir(dir):
    if not os.path.exists(dir):
        os.makedirs(dir)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input_dir', type=str, default=os.getcwd(), help='Path to the directory containing the excel files to input.')
    parser.add_argument('-o', '--output_dir', type=str, default=os.getcwd(), help='Path to the directory for output pdf files.')
    parser.add_argument('-d', '--divide', action='store_true', default=False, help='Divide mode, which will convert each worksheet into a separated pdf file.')
    parser.add_argument('-s', '--sheets', type=int, nargs="+", default=[], help='When divide mode is enabled, which sheets should be converted for each excel file. Example: -s 1 2 3 . Leave it blank if you want to convert all sheets.')
    parser.add_argument('-r', '--orientation', type=int, default=1, help='Orientation of pdf file. 1: Portrait 2: Landscape')
    parser.add_argument('-z', '--zoom', type=int, default=0, help='Zoom sheets to a single page in specific direction. 0: Disable 1: Zoom Tall 2: Zoom Wide')
    parser.add_argument('-m', '--merge', action='store_true', default=False, help='Automatically merge all converted pdf files into a single pdf file')

    args = parser.parse_args()
    input_dir = args.input_dir
    output_dir = args.output_dir
    divide = args.divide
    sheets_to_convert = args.sheets
    orientation = args.orientation
    merge = args.merge
    zoom = args.zoom

    mkdir(output_dir)

    files_path = []
    for root, dirs, files in os.walk(args.input_dir):
        files_path += [os.path.join(root, f) for f in files if f.split('.')[-1].upper() in ["XLS", "XLSX"]]

    xl = DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = 0

    pdf_merger = PdfMerger()

    atexit.register(xl.Quit)

    print("Converting ...")

    for input_path in tqdm(files_path):
        workbook = None
        try:
            workbook = xl.Workbooks.Open(input_path)
            for sheet in workbook.Worksheets:
                sheet.PageSetup.Orientation = orientation
                if zoom == 1:
                    sheet.PageSetup.Zoom = False
                    sheet.PageSetup.FitToPagesTall = 1
                elif zoom == 2:
                    sheet.PageSetup.Zoom = False
                    sheet.PageSetup.FitToPagesWide = 1
            if divide:  # 启用拆分模式
                if sheets_to_convert:  # 指定了需要转换的工作表
                    for sheet in sheets_to_convert:
                        try:
                            output_file = os.path.join(output_dir, f"{os.path.basename(input_path)}_Sheet{sheet}.pdf")
                            worksheet = workbook.Worksheets[sheet-1]
                            worksheet.ExportAsFixedFormat(0, output_file)
                            if merge:
                                pdf_merger.append(output_file)
                        except Exception as e:
                            print(e)
                            pass
                else:  # 转换所有工作表
                    for i, sheet in enumerate(workbook.Worksheets):
                        output_file = os.path.join(output_dir, f"{os.path.basename(input_path)}_Sheet{i+1}.pdf")
                        sheet.ExportAsFixedFormat(0, output_file)
                        if merge:
                            pdf_merger.append(output_file)
            else:  # 不启用拆分模式
                output_file = os.path.join(output_dir, f"{os.path.basename(input_path)}.pdf")
                workbook.ExportAsFixedFormat(0, output_file)
                if merge:
                    pdf_merger.append(output_file)
        except Exception as e:
            print(e)
        finally:
            workbook.Close(False)

    if merge:
        print("Merging ...")
        pdf_merger.write(os.path.join(output_dir, "merged.pdf"))

    xl.Quit()
