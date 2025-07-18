from config import Config
from counting import CountReportFromWorklist
from data_pipeline import DataPipeline
from dotenv import load_dotenv
from download_file import DownloadFile
from excel_manipulation import ExcelModule
from file_module import FileModule
from logger_config import setup_workilist_logger
import os
from report_dictionary import WorklistDictionary
from run_macro import RunMacro
import time
from validation import ValidationOfFiles

load_dotenv()

logger = setup_workilist_logger(__name__)

def Worklist() -> None:
    """Worklist Automation"""
    try:
        username: str = os.getenv("SHAREPOINT_USERNAME")
        password: str = os.getenv("SHAREPOINT_PASSWORD")
        site_url: str = os.getenv("SHAREPOINT_URL")


        file_management = FileModule()
        pipeline = DataPipeline()
        worklist = WorklistDictionary()
        excel_module = ExcelModule()
        sites = ['Scott', "Wheat Ridge", "Orlando", "Dayton"]
        sharepoint_file = DownloadFile(username, password, site_url)


        for site in sites:
            # -------------------------------------------------Scott Automation----------------------------------------------------------------------------------------------------------------------#



            # if site == "Scott":
            #     logger.info(f"Automating the worklist report. Site: {site}")
            #     files:list[str] = file_management.get_data_from_folder(Config.Scott().raw_file_path)

            #     logger.info(f"Validating Files")
            #     validated_files:list[str] =  ValidationOfFiles.validate_worklist_file(files, Config.Scott().worklist_report_names)

            #     if len(validated_files) != 0:

            #         logger.info("Running Macro Worklist (Scott)...")
            #         RunMacro(validated_files, Config.macro_path, Config.macro_hotkey_worklist)

            #         logger.info("Cleaning the raw files. (Data Pipeline Running...)")
            #         pipeline.pipeline_for_worklist(validated_files, Config.Scott().yesterday_report_path, Config.Scott().qa_samples, worklist.Scott, Config.Scott().counting_data)

            #         logger.info("Inserting the data into the template...")
            #         excel_module.insert_data_from_template(validated_files, Config.Scott().template_file_path, Config.Scott().save_name_file, 1, worklist.Scott)


            #         logger.info("Adding count data on daily Data")
            #         CountReportFromWorklist(Config.Scott().counting_data, Config.Scott().save_name_file, site)


            #         logger.info(f"Automation Success (Worklist). Done automating site {site}. Please check the files for validation.")
            
            #     elif len(files) == 0:
            #         logger.warning(f"No Worklist Files found in the folder: {site}")



            # --------------------------------------------------Wheat Ridge Automation--------------------------------------------------------------------------------------------------------------------#


            # if site == "Wheat Ridge":


            #     logger.info(f"Automating the worklist report. Site: {site}")
            #     files:list[str] = file_management.get_data_from_folder(Config.Wheat_Ridge().raw_file_path)

                
            #     logger.info(f"Validating Files")
            #     validated_files:list[str] =  ValidationOfFiles.validate_worklist_file(files, Config.Wheat_Ridge().worklist_report_names)

            #     if len(validated_files) != 0:

            #         logger.info("Running Macro Worklist (Wheat Ridge)...")
            #         RunMacro(validated_files, Config.macro_path, Config.macro_hotkey_worklist)

                    
            #         logger.info("Cleaning the raw files. (Data Pipeline Running...)")
            #         pipeline.pipeline_for_worklist(validated_files, Config.Wheat_Ridge().yesterday_report_path, Config.Wheat_Ridge().qa_samples, worklist.Wheat_Ridge, Config.Wheat_Ridge().counting_data)

                    
            #         logger.info("Inserting the data into the template...")
            #         excel_module.insert_data_from_template(validated_files, Config.Wheat_Ridge().template_file_path, Config.Wheat_Ridge().save_name_file, 1, worklist.Wheat_Ridge)

            #         logger.info("Adding count data on daily Data")
            #         CountReportFromWorklist(Config.Wheat_Ridge().counting_data, Config.Wheat_Ridge().save_name_file, site)

            #         logger.info(f"Fixing Format in Worklist ({site})")
            #         RunMacro(Config.Wheat_Ridge().reported_report, Config.macro_path_daily_task, Config.macro_hotkey_fixed_wheat_format)

            #         logger.info(f"Automation Success. Done automating site {site}. Please checkt the files for validation.")
                                
            #     elif len(files) == 0:
            #         logger.warning(f"No Worklist Files files found in the folder: {site}")


            # -----------------------------------------------------Orlando Automation-----------------------------------------------------------------------------------------------------------------#


            # if site == "Orlando":
            #     logger.info(f"Automating the worklist report. Site: {site}")
            #     files:list[str] = file_management.get_data_from_folder(Config.Orlando().raw_file_path)

            #     logger.info(f"Validating Files")
            #     validated_files:list[str] =  ValidationOfFiles.validate_worklist_file(files, Config.Orlando().worklist_report_names)

            #     if len(validated_files) != 0:
                    

            #         logger.info("Running Macro Worklist (Orlando)...")
            #         RunMacro(validated_files, Config.macro_path, Config.macro_hotkey_worklist)

            #         logger.info("Cleaning the raw files. (Data Pipeline Running...)")
            #         pipeline.pipeline_for_worklist(validated_files, Config.Orlando().yesterday_report_path, Config.Orlando().qa_samples, worklist.Orlando, Config.Orlando().counting_data)

            #         logger.info("Inserting the data into the template...")
            #         excel_module.insert_data_from_template(validated_files, Config.Orlando().template_file_path, Config.Orlando().save_name_file, 1, worklist.Orlando)

            #         # logger.info("Adding count data on daily Data")
            #         # CountReportFromWorklist(Config.Orlando().counting_data, Config.Orlando().save_name_file, site)

            #         logger.info(f"Fixing Format in Worklist ({site})")
            #         RunMacro(Config.Orlando().reported_report, Config.macro_path_daily_task, Config.macro_hotkey_fixed_orlando_format)

            #         logger.info(f"Automation Success (Worklist). Done automating site {site}. Please check the files for validation.")
                

            
            #     elif len(files) == 0:
            #         logger.warning(f"No Worklist Files found in the folder: {site}")
                


            # ---------------------------------------------------------Dayton Automation----------------------------------------------------------------------------------------------------------------#



            if site == "Dayton":
                logger.info(f"Automating the worklist report. Site: {site}")
                files:list[str] = file_management.get_data_from_folder(Config.Dayton().raw_file_path)

                logger.info(f"Validating Files")
                validated_files:list[str] =  ValidationOfFiles.validate_worklist_file(files, Config.Dayton().worklist_report_names)

                if len(validated_files) != 0:
                    

                    logger.info("Running Macro Worklist (Dayton)...")
                    RunMacro(validated_files, Config.macro_path, Config.macro_hotkey_worklist)

                    logger.info("Cleaning the raw files. (Data Pipeline Running...)")
                    pipeline.pipeline_for_worklist(validated_files, Config.Dayton().yesterday_report_path, Config.Dayton().qa_samples, worklist.Dayton, Config.Dayton().counting_data, site)

                    logger.info("Inserting the data into the template...")
                    excel_module.insert_data_from_template(validated_files, Config.Dayton().template_file_path, Config.Dayton().save_name_file, 5, worklist.Dayton)


                    logger.info("Adding count data on daily Data")
                    CountReportFromWorklist(Config.Dayton().counting_data, Config.Dayton().save_name_file, site)

                    logger.info(f"Fixing Format in Worklist ({site})")
                    RunMacro(Config.Dayton().reported_report, Config.macro_path_daily_task, Config.macro_hotkey_fixed_dayton_format)
                    

                    logger.info(f"Automation Success (Worklist). Done automating site {site}. Please check the files for validation.")
                

            
                elif len(files) == 0:
                    logger.warning(f"No Worklist Files found in the folder: {site}")
            
            # -------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
   


    except Exception as e:
        logger.critical(f"Automation failed in worklies, Error: {e}", exc_info=True)

        


if __name__ == "__main__":
    Worklist()



    
    # new_files: list[str] = file_management.move_file("\\data\\Scott\\raw", "data\\Scott\\history-raw", isEndWith=True, endswith = ".xls")

