from config import Config
from counting import CountReportFromWorklist
from data_pipeline import DataPipeline
from excel_manipulation import ExcelModule
from file_module import FileModule
from logger_config import setup_workilist_logger
from report_dictionary import WorklistDictionary
from run_macro import RunMacro
from validation import ValidationOfFiles




logger = setup_workilist_logger(__name__)

def Worklist() -> None:
    """Worklist Automation"""
    try:
        file_management = FileModule()
        pipeline = DataPipeline()
        worklist = WorklistDictionary()
        excel_module = ExcelModule()
        sites = ['Scott', "Wheat Ridge"]


        for site in sites:

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
            #         #CountReportFromWorklist(Config.Scott().counting_data, "test_automation.xlsx", site)

            #         logger.info(f"Fixing Format in Worklist ({site})")
            #         RunMacro(Config.Scott().reported_report, Config.macro_path_daily_task, Config.macro_hotkey_fixed_scott_format)

            #         logger.info(f"Automation Success (Worklist). Done automating site {site}. Please check the files for validation.")
            
            #     elif len(files) == 0:
            #         logger.warning(f"No Worklist Files found in the folder: {site}")

            if site == "Wheat Ridge":


                logger.info(f"Automating the worklist report. Site: {site}")
                files:list[str] = file_management.get_data_from_folder(Config.Wheat_Ridge().raw_file_path)

                
                logger.info(f"Validating Files")
                validated_files:list[str] =  ValidationOfFiles.validate_worklist_file(files, Config.Wheat_Ridge().worklist_report_names)

                if len(validated_files) != 0:

                    logger.info("Running Macro Worklist (Wheat Ridge)...")
                    RunMacro(validated_files, Config.macro_path, Config.macro_hotkey_worklist)

                    
                    logger.info("Cleaning the raw files. (Data Pipeline Running...)")
                    pipeline.pipeline_for_worklist(validated_files, Config.Wheat_Ridge().yesterday_report_path, Config.Wheat_Ridge().qa_samples, worklist.Wheat_Ridge, Config.Wheat_Ridge().counting_data)

                    
                    logger.info("Inserting the data into the template...")
                    excel_module.insert_data_from_template(validated_files, Config.Wheat_Ridge().template_file_path, Config.Wheat_Ridge().save_name_file, 1, worklist.Wheat_Ridge)

                    logger.info("Adding count data on daily Data")
                    CountReportFromWorklist(Config.Wheat_Ridge().counting_data, Config.Wheat_Ridge().save_name_file, site)

                    logger.info(f"Fixing Format in Worklist ({site})")
                    RunMacro(Config.Wheat_Ridge().reported_report, Config.macro_path_daily_task, Config.macro_hotkey_fixed_wheat_format)

                    logger.info(f"Automation Success. Done automating site {site}. Please checkt the files for validation.")
                                
                elif len(files) == 0:
                    logger.warning(f"No Worklist Files files found in the folder: {site}")



    except Exception as e:
        logger.critical(f"Automation failed in worklies, Error: {e}", exc_info=True)

        


if __name__ == "__main__":
    Worklist()



    
    # new_files: list[str] = file_management.move_file("\\data\\Scott\\raw", "data\\Scott\\history-raw", isEndWith=True, endswith = ".xls")

