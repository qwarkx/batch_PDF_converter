import argparse

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
    parser.add_argument('--folder_path', type=int, default=folder_path)
    parser.add_argument('--rtc', type=str, default=rtc)
    args = parser.parse_args()


    # Main code which ia run
    run_program(args)

def run_program(args):

    # Initialise code and arguments

    file_name= args.folder_path
    time_rtc = args.rtc

    print(args)


    try:
        print("generating PDF from word files")

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
