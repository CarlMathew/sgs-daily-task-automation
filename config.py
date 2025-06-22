from datetime import datetime
from file_module import FileModule



# Configuration Files for every site

class Config:


    macro_path:str = "Copy of PERSONAL.xlsb"
    macro_path_daily_task:str = "DailyTaskFixFormat.xlsm"

    macro_hotkey_worklist:str = "Backlog"
    macro_hotkey_rush: str = "Rush"

    macro_hotkey_fixed_scott_format = "WorklistFixedDateFormatScott"
    macro_hotkey_fixed_wheat_format = "WorklistFixedDateFormatWheatRidge"
    
    macro_hotkey_scott_rush:str = "RushScott"
    macro_hotkey_wheat_ridge_rush:str = "RushWheatRidge"




    class Scott(FileModule):
        def __init__(self):
            super().__init__()
            self.qa_samples:str = "ALLA"
            self.raw_file_path:str = "\\data\\Scott\\raw"
            self.template_file_path:str = "template\\Scott Late Report Temp v3.xlsx"
            self.yesterday_report = "\\data\\Scott\\yesterday-report\\"
            self.yesterday_report_path:str = self.get_data_from_folder(self.yesterday_report)[-1]
            self.counting_data: str = "Daily Count (Report)\\Scott\\WorkList_Count (Scott).csv"
            self.rush_counting_data: str = "Daily Count (Report)\\Scott\\Rush_Count (Scott).csv"
            self.save_name_file:str = f"data\\Scott\\report\\Scott_Late_{datetime.now().strftime("%y%m%d")}.xlsx"
            self.worklist_report_names: list[str]= ["aaall", "ext-nd", "gcs", "gcv", "gnallnd", "hgall", "mss", "msvoa"]
            self.reported_report = self.get_data_from_folder("\\data\\Scott\\report\\")



            self.archived_path: str = ""
    
    class Wheat_Ridge(FileModule):
        def __init__(self):
            super().__init__()
            self.qa_samples:str = "ALMS"
            self.raw_file_path:str = "\\data\\Wheat Ridge\\raw"
            self.template_file_path:str = "template\\WheatRidge Late Report Temp v4.xlsx"
            self.yesterday_report = "\\data\\Wheat Ridge\\yesterday-report"
            self.yesterday_report_path:str = self.get_data_from_folder(self.yesterday_report)[-1]
            self.counting_data: str = "Daily Count (Report)\\Wheat Ridge\\WorkList_Count (Wheat Ridge).csv"
            self.rush_counting_data: str = "Daily Count (Report)\\Scott\\Rush_Count (Wheat Ridge).csv"
            self.save_name_file:str = f"data\\Wheat Ridge\\report\\Wheat_Ridge_Late_{datetime.now().strftime("%y%m%d")}.xlsx"
            self.worklist_report_names: list[str]= ["aaallnd", "orgprep", "gcs", "gcvoa", "gnall", "hgall", "mss", "msvoa"]
            self.reported_report = self.get_data_from_folder("\\data\\Wheat Ridge\\report\\")


            self.archived_path: str = ""
