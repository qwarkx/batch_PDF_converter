

import os
import argparse

from win32com import client
from pathlib import Path

from datetime import datetime
import time

def main():


    parser = argparse.ArgumentParser()

    now = datetime.now()
    r_date = now.strftime("%Y-%m-%d")
    r_time = now.strftime("%H_%M_%S")
    rtc = (r_date + '_' + r_time).replace(" ","_").replace(":","_").replace("/","_")

    # ---- Set File Name Here

    folder_path = "CTTpymat_" + "SIG_Tester" + "_V800"

    # NOTE: Select waht to read, device[1] or sensor[0]
    parser.add_argument('--folder_path', type=str, default=folder_path)
    parser.add_argument('--rtc', type=str, default=rtc)
    args = parser.parse_args()


    # Main code which ia run
    run_program(args)



def convert_docx_to_pdf(docx_path, pdf_path):
    word = client.Dispatch('Word.Application')
    doc = word.Documents.Open(docx_path)
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    word.Quit()

# Example usage
# convert_docx_to_pdf('example.docx', 'example.pdf')

def batch_convert_docx_to_pdf(file_list, output_folder = None):

    if not file_list:
        print("No files to convert.")
        return

    output_folder = os.path.join(os.path.dirname(file_list[0]), "PDF")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    try:
        print("Converting docx to PDF:")

        for docx_path in file_list:
            pdf_path = os.path.join(output_folder, f'{os.path.splitext(os.path.basename(docx_path))[0]}.pdf')
            print(' * ' + pdf_path)
            convert_docx_to_pdf(docx_path, pdf_path)
    except Exception as e:
        print("Something went wrong in printing, see:  ", e)
    print('Finished converting the files')

def list_files(folder_path, extensions=['.docx']):
    files = []
    try:
        for root, _, filenames in os.walk(folder_path):
            for filename in filenames:
                if any(filename.endswith(ext) for ext in extensions):
                    file_path = os.path.join(root, filename)
                    print(file_path)
                    files.append(file_path)
    except Exception as e:
        print('Files could not be located, see:', e)
    return files

def run_program(args):

    # Initialise code and arguments

    file_name= args.folder_path
    time_rtc = args.rtc

    print(args)


    try:
        print("Get file paths")
        file_paths = 'D:\[Part 100 ] - Device documents V65\[Part 016 ] - Technical Files\SIQ PDF files\\testing'

        files = list_files(file_paths)
        batch_convert_docx_to_pdf(files)

        print("generating PDF from word files", files)
    except Exception as e:
        print('Something happened while exporting, see:  ', e)
    try:
        print("Exporting the file list of Technical Files sending")
    except Exception as e:
        print('Something happened while making the TF file list, see:  ', e)



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
