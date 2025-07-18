from datetime import datetime
from file_module import FileModule



# Configuration Files for every site

class Config:


    macro_path:str = "Copy of PERSONAL.xlsb"
    macro_path_daily_task:str = "DailyTaskFixFormat.xlsm"

    macro_hotkey_worklist:str = "Backlog"
    macro_hotkey_rush: str = "Rush"
    macro_hotkey_late:str = "LateReport"

    macro_hotkey_fixed_scott_format = "WorklistFixedDateFormatScott"
    macro_hotkey_fixed_wheat_format = "WorklistFixedDateFormatWheatRidge"
    macro_hotkey_fixed_orlando_format = "WorklistFixedDateFormatOrlando"
    macro_hotkey_fixed_dayton_format = "WorklistFixedDateFormatDayton"

    macro_sort_due_date_scott = "SortDueDateScott"
    macro_sort_due_date_wheat_ridge = "SortDueDateWheatRidge"
    
    macro_hotkey_scott_rush:str = "RushScott"
    macro_hotkey_wheat_ridge_rush:str = "RushWheatRidge"
    macro_hotkey_orlando_rush:str = "RushOrlando"
    macro_hotkey_dayton_rush:str = "RushDayton"

    macro_hoteky_scott_late = "LateScott" 
    macro_hoteky_wheat_ridge_late:str = "LateWheatRidge"
    macro_hoteky_orlando_late:str = "LateOrlando"
    macro_hoteky_dayton_late:str = "LateDayton"



    class Scott(FileModule):
        def __init__(self):
            super().__init__()
            self.qa_samples:str = "ALLA"
            self.raw_file_path:str = "\\data\\Scott\\raw"
            self.raw_file_path_V2:str = "data\\Scott\\raw"
            self.template_file_path:str = "template\\Scott Late Report Temp v3.xlsx"
            self.yesterday_report = "\\data\\Scott\\yesterday-report\\"
            self.yesterday_report_V2 = "data\\Scott\\yesterday-report\\"
            self.yesterday_report_path:str = self.get_data_from_folder(self.yesterday_report)[-1]
            self.counting_data: str = "Daily Count (Report)\\Scott\\WorkList_Count (Scott).csv"
            self.rush_counting_data: str = "Daily Count (Report)\\Scott\\Rush_Count (Scott).csv"
            self.save_name_file:str = f"data\\Scott\\report\\Scott_Late_{datetime.now().strftime("%y%m%d")} remove.xlsx"
            self.worklist_report_names: list[str]= ["aaall", "ext-nd", "gcs", "gcv", "gnallnd", "hgall", "mss", "msvoa"]
            self.reported_report = self.get_data_from_folder("\\data\\Scott\\report\\")
            self.sharepoint_url = "/sites/ph-ehs-ofs/EHS/GBS EHS US/8 References/US QA Coordinator/Daily Dashboard and Late Report/Scott Dashboard and Late Report/Daily Late Report"



            self.archived_path: str = ""
    
    class Wheat_Ridge(FileModule):
        def __init__(self):
            super().__init__()
            self.qa_samples:str = "ALMS"
            self.raw_file_path:str = "\\data\\Wheat Ridge\\raw"
            self.raw_file_path_V2:str = "data\\Wheat Ridge\\raw"
            self.template_file_path:str = "template\\WheatRidge Late Report Temp v4.xlsx"
            self.yesterday_report = "\\data\\Wheat Ridge\\yesterday-report"
            self.yesterday_report_V2 = "data\\Wheat Ridge\\yesterday-report\\"
            self.yesterday_report_path:str = self.get_data_from_folder(self.yesterday_report)[-1]
            self.counting_data: str = "Daily Count (Report)\\Wheat Ridge\\WorkList_Count (Wheat Ridge).csv"
            self.rush_counting_data: str = "Daily Count (Report)\\Wheat Ridge\\Rush_Count (Wheat Ridge).csv"
            self.save_name_file:str = f"data\\Wheat Ridge\\report\\Wheat_Ridge_Late_{datetime.now().strftime("%y%m%d")} remove.xlsx"
            self.worklist_report_names: list[str]= ["aaallnd", "orgprep", "gcs", "gcvoa", "gnall", "hgall", "mss", "msvoa"]
            self.reported_report = self.get_data_from_folder("\\data\\Wheat Ridge\\report\\")
            self.sharepoint_url = "/sites/ph-ehs-ofs/EHS/GBS EHS US/8 References/US QA Coordinator/Daily Dashboard and Late Report/WheatRidge Dashboard and Late Report/Daily Late Report"


            self.archived_path: str = ""

    class Orlando(FileModule):
        def __init__(self):
            super().__init__()
            self.qa_samples:str = "ALSE"
            self.raw_file_path:str = "\\data\\Orlando\\raw"
            self.raw_file_path_V2:str = "data\\Orlando\\raw"
            self.template_file_path:str = "template\\Orlando Late Report Temp v2.xlsx"
            self.yesterday_report = "\\data\\Orlando\\yesterday-report"
            self.yesterday_report_V2 = "data\\Orlando\\yesterday-report\\"
            self.yesterday_report_path:str = self.get_data_from_folder(self.yesterday_report)[-1]
            self.counting_data: str = "Daily Count (Report)\\Orlando\\WorkList_Count (Orlando).csv"
            self.rush_counting_data: str = "Daily Count (Report)\\Orlando\\Rush_Count (Orlando).csv"
            self.save_name_file:str = f"data\\Orlando\\report\\Orlando_Late_{datetime.now().strftime("%y%m%d")} remove.xlsx"
            self.worklist_report_names: list[str]= ["aaallnd", "extnd", "gcmsnd", "gcnd", "gcvnd", "gnallnd", "hg", "lc-qqq", "lc-qqqprep", "msvnd"]
            self.reported_report = self.get_data_from_folder("\\data\\Orlando\\report\\")
            self.sharepoint_url = "/sites/ph-ehs-ofs/EHS/GBS EHS US/8 References/US QA Coordinator/Daily Dashboard and Late Report/WheatRidge Dashboard and Late Report/Daily Late Report"

            self.archived_path: str = ""


    class Dayton(FileModule):
        def __init__(self):
            super().__init__()
            self.qa_samples:str = "ALNJ"
            self.raw_file_path:str = "\\data\\Dayton\\raw"
            self.raw_file_path_V2:str = "data\\Dayton\\raw"
            self.template_file_path:str = "template\\Dayton Late Report Temp v4.xlsx"
            self.yesterday_report = "\\data\\Dayton\\yesterday-report"
            self.yesterday_report_V2 = "data\\Dayton\\yesterday-report\\"
            self.yesterday_report_path:str = self.get_data_from_folder(self.yesterday_report)[-1]
            self.counting_data: str = "Daily Count (Report)\\Dayton\\WorkList_Count (Dayton).csv"
            self.rush_counting_data: str = "Daily Count (Report)\\Dayton\\Rush_Count (Dayton).csv"
            self.save_name_file:str = f"data\\Dayton\\report\\Dayton_Late_{datetime.now().strftime("%y%m%d")} remove.xlsx"
            self.worklist_report_names: list[str]= ["aaallnd", "extlcms", "ext-nd", "gcair", "gcs", "gcvoa", "gnallnd", "hgall", "lcmspfas", "msair", "mss", "msvoa"]
            self.reported_report = self.get_data_from_folder("\\data\\Dayton\\report\\")
            self.sharepoint_url = "/sites/ph-ehs-ofs/EHS/GBS EHS US/8 References/US QA Coordinator/Daily Dashboard and Late Report/WheatRidge Dashboard and Late Report/Daily Late Report"

            self.archived_path: str = ""
