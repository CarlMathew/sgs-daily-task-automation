from datetime import datetime
import openpyxl.workbook
import pandas as pd
from openpyxl import load_workbook

def AddCellValueOnDataFrame(df:pd.DataFrame, report_type:str) -> pd.DataFrame:
    """add the cell value in the dataframe that can be used to edit the excel"""

    if report_type.title() == "Scott":
        df.loc[df["ReportType"] == "METALS", "Cell_Value"] = "B11"
        df.loc[df["ReportType"] == "ORGPREP", "Cell_Value"] = "B12"
        df.loc[df["ReportType"] == "GCSEMI", "Cell_Value"] = "B7" 
        df.loc[df["ReportType"] == "GCVOA", "Cell_Value"] = "B9"
        df.loc[df["ReportType"] == "GENCHEM", "Cell_Value"] ="B10"
        # df.loc[df["Type"] == "HG", "Cell_Value"] = ""
        df.loc[df["ReportType"] == "MSSEMI", "Cell_Value"] = "B6"
        df.loc[df["ReportType"] == "MSVOA", "Cell_Value"] = "B8"

    elif report_type.title() == "Wheat Ridge":
        df.loc[df["ReportType"] == "METALS", "Cell_Value"] = "B11"
        df.loc[df["ReportType"] == "ORGPREP", "Cell_Value"] = "B12"
        df.loc[df["ReportType"] == "GCSEMI", "Cell_Value"] = "B7" 
        df.loc[df["ReportType"] == "GCVOA", "Cell_Value"] = "B9"
        df.loc[df["ReportType"] == "GENCHEM", "Cell_Value"] ="B10"
        # df.loc[df["Type"] == "HG", "Cell_Value"] = ""
        df.loc[df["ReportType"] == "MSSEMI", "Cell_Value"] = "B6"
        df.loc[df["ReportType"] == "MSVOA", "Cell_Value"] = "B8"



    return df

def AddDataToReport(df:pd.DataFrame, ExcelPath) -> None:
    """Add the count that comes from data cleaning"""
    data:list[dict] = df.to_dict(orient="records")

    wb: openpyxl.workbook = load_workbook(ExcelPath)
    count_sheet = wb["COUNT"]

    # Add Data on the count tab
    for datum in data:
        try:
            cell_position: str = datum["Cell_Value"]
            cell_value: int = datum["LateCount"]
            count_sheet[cell_position] = cell_value
        except Exception as e:
            pass

    wb.save(ExcelPath)
    wb.close()


def CountReportFromWorklist(report_count_path:str, ExcelFilePath="", report:str="") -> None:
    """ Fetch data count from csv then put it on the report"""
    now = datetime.now().date()
    
    df = pd.read_csv(report_count_path)
        
    df = df.loc[df["Date"] == str(now)]

    df = AddCellValueOnDataFrame(df, report)

    AddDataToReport(df, ExcelFilePath)


# if __name__ == "__main__":
#     CountReportFromWorklist("Daily Count (Report)\\WorkList_Count (Scott).csv", "test_automation.xlsx", "Scott")