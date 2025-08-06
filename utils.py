from config import Config
from datetime import datetime, timedelta
import re
import pandas as pd

class Utils:
    # Improvement: Also add the holiday in US
    
    @staticmethod
    def get_total_workday(days: int) -> int:
        """
            Fetch the total if weekends are not included on the days
            
            Args:
                days: int -> working days
            
            Returns:
                int: -> return how many the workday is without including the sat and sunday
        """
        for i in range(days):
            current_day = datetime.now() - timedelta(days=i)
            day_of_week = current_day.strftime("%A")

            if day_of_week in ["Saturday", "Sunday"]:
                days += 1 
    
        return days
    

    @staticmethod
    def get_late_tab_name_in_excel(excel_file: str) -> str:
        """
            List all the tab name inside the excel and return the late tab string

            Args:
                excel_file: str -> excel file path
            
            
            Returns:
                str: Name of the late tab inside the excel
        """


        df = pd.ExcelFile(excel_file)

        sheet_names: list = df.sheet_names
        
        pattern_re: str = re.compile(r"_Late_\d+")

        for sheet in sheet_names:
            if pattern_re.search(sheet):
                return sheet


    @staticmethod
    def get_repgen_columns_in_excel(df: pd.DataFrame) -> str:

        """
            Get the column name of repgen in late report

            Args:
                excel_file: (str) -> excel file path
            
            Returns:
                str: return the name of the comments column in repgen tab
        """

        column_list: list = list(df.columns)
        
        pattern_re: re = re.compile(r"\b(Repgen)\b", re.IGNORECASE)
        

        for column in column_list:
            if pattern_re.search(column):
                return column



if __name__ == "__main__":
    yesterday_report = Config.Orlando().yesterday_report_path
    utility = Utils()

    repgen: str =  utility.get_repgen_columns_in_excel(yesterday_report)
    print(repgen)

