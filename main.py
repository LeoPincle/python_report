import sys
import win32com.client
import os
import glob


from ITSM_Excel.Excel_Generate.writeFileExcel import WriteFileExcel
from ITSM_Util.inputUtil import InputUtil
from ITSM_PPT.writeFilePPT import WriteFilePPT
from ITSM_Excel.Excel_FetchData.fetchData import FetchData

if __name__ == '__main__':
    print("HP Weekly ITSM Report\n")

    """Create required output directories"""
    if not os.path.exists("Intermediary"):
        os.makedirs("Intermediary")

    if not os.path.exists("Output"):
        os.makedirs("Output")

    """Create input util object to take date and project input from user"""

    input_util = InputUtil()
    input_res = input_util
    print("\tProcessing Excel Data...")

    res = dict(map(lambda i,j : (i,j) , input_util.get_project_keys(),input_util.get_project_names()))

    for value in res:

        fetch_data = FetchData(input_util.get_startDate(), input_util.get_endDate(), value)

        obj_excel = WriteFileExcel(res[value], input_util.get_current_week_date_range_file_name(), fetch_data)

        try:
            obj_excel.write()
        except IndexError as ie:
            print("\nUnexpected Error in writing Excel : ", ie)
        except ValueError as ve:
            print("\nUnexpected Error in writing Excel : ", ve)
        except Exception as exc:
            print("Error", exc,
                  "\nUnexpected Error :  Check if data for Project entered is present in the Excel dump  \n")
            print("Press Enter to close the window\n")
            input()
            sys.exit()
        print("\tExcel data processed.\n")

    """Write data to PPT"""
    print("\tWriting data to PPT...")


    obj_ppt = WriteFilePPT(input_util)
    try:
        obj_ppt.write()
    except IndexError as ie:
        print("\nUnexpected Error in writing data to PPT : ", ie)
    except ValueError as ve:
        print("\nUnexpected Error in writing data to PPT : ", ve)
    except Exception as exc:
        print("Error : ", exc, "\nUnexpected Error in writing data to PPT")

    if len(input_util.get_project_names()) >1:

        pres = glob.glob("Output/*.pptx")
        output_path = "Output/GCO - Weekly ITSM Report - " + input_util.get_current_week_date_range_file_name() + ".pptx"

        def merge_presentations(presentations, path):
            ppt_instance = win32com.client.Dispatch('PowerPoint.Application')
            prs = ppt_instance.Presentations.open(os.path.abspath(presentations[0]), True, False, False)

            for i in range(1, len(presentations)):
                prs.Slides.InsertFromFile(os.path.abspath(presentations[i]), prs.Slides.Count)

            prs.SaveAs(os.path.abspath(path))
            prs.Close()


        merge_presentations(pres, output_path)

    print("\tData written to PPT.\n")
    print("Task Completed\n")
    print("\n Press Enter to close the window\n")
    input("")