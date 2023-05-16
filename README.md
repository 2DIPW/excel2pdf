# Excel2PDF
A python script for batch converting Excel files to PDF files, which support specifying paper orientation, scaling oversized worksheets to one page, and merging all outputs to one PDF file.
## How to use
Use [excel2pdf.py](https://github.com/2DIPW/excel2pdf/blob/master/excel2pdf.py)

- `-i` | `--input_dir`: Path to the directory containing the excel files to input.
- `-o` | `--output_dir`: Path to the directory for output pdf files.

  > If -p and -f are not specified, the default value is current directory.
- `-d` | `--divide`: Divide mode, which will convert each worksheet into a separated pdf file.
- `-s` | `--sheets`: When divide mode is enabled, which sheets should be converted for each excel file. *Example: -s 1 2 3.* Leave it blank if you want to convert all sheets
- `-r` | `--orientation`: Orientation of pdf file. **1: Portrait(Default)** 2: Landscape
- `-m` | `--merge`: Automatically merge all converted pdf files into a single pdf file.
- `-z` | `--zoom`: Zoom excel file to a single page. **0: Disable(Dafault)** 1: Zoom Tall 2: Zoom Wide

## Example usage
- If you want to convert all excel files in the current directory to pdf and merge them into one pdf file.
    ```shell
    python excel2pdf.py -m
    ```
- If you want to convert the first and second worksheets of all excel files in the current directory into separate pdfs, specify the output orientation as landscape, and each worksheet is horizontally scaled to one page.
    ```shell
    python excel2pdf.py -d -s 1 2 -r 2 -z 1
    ```