from config import Config
from counting import CountReportFromWorklist
from data_pipeline import DataPipeline
from excel_manipulation import ExcelModule
from file_module import FileModule
from logger_config import setup_rush_logger
from report_dictionary import WorklistDictionary
from run_macro import RunMacro
from validation import ValidationOfFiles


logger = setup_rush_logger(__name__)

def Rush() -> None:
    """Rush Automation"""


    try:
        file_management = FileModule()
        pipeline = DataPipeline()
        # worklist = WorklistDictionary()
        excel_module = ExcelModule()
        sites = ['Scott', "Wheat Ridge"]


        for site in sites:
            # if site == "Scott":
            #     logger.info(f"Automating the rush report. Site: {site}")
            #     files:list[str] = file_management.get_data_from_folder(Config.Scott().raw_file_path)

                
            #     logger.info(f"Validating Files")
            #     validated_files:list[str] =  ValidationOfFiles.validate_rush_file(files)

            #     if len(validated_files) == 1:
                    
            #         logger.info("Running Macro Rush...")
            #         RunMacro(validated_files, Config.macro_path, Config.macro_hotkey_rush)

            #         logger.info("Cleaning the rush raw files. (Data Pipeline is running...)")
            #         pipeline.pipeline_for_rush(validated_files, Config.Scott().yesterday_report_path, Config.Scott().qa_samples, site, Config.Scott().rush_counting_data)


            #         logger.info("Inserting the data into the template...")
            #         excel_module.insert_data_from_template(validated_files, Config.Scott().save_name_file, Config.Scott().save_name_file, 2, "RUSH_PRIORITY")
                                        
            #         logger.info("Running the macro for the xloopup rush")
            #         RunMacro(Config.Scott().reported_report, Config.macro_path_daily_task, Config.macro_hotkey_scott_rush)

            #         logger.info(f"Automation Success (Rush). Done automating site {site}. Please checkt the files for validation.")


            #     elif len(validated_files) == 0:
            #         logger.warning(f"No Rush Files found in the folder: {site}")

            if site == "Wheat Ridge":
                logger.info(f"Automating the rush report. Site: {site}")
                files:list[str] = file_management.get_data_from_folder(Config.Wheat_Ridge().raw_file_path)



                logger.info(f"Validating Files")
                validated_files:list[str] =  ValidationOfFiles.validate_rush_file(files)


                if len(validated_files) == 1:

                    logger.info("Running Macro Rush...")
                    RunMacro(validated_files, Config.macro_path, Config.macro_hotkey_rush)

                    logger.info("Cleaning the rush raw files. (Data Pipeline is running...)")
                    pipeline.pipeline_for_rush(validated_files, Config.Wheat_Ridge().yesterday_report_path, Config.Wheat_Ridge().qa_samples, site, Config.Wheat_Ridge().rush_counting_data)


                    logger.info("Inserting the data into the template...")
                    excel_module.insert_data_from_template(validated_files, Config.Wheat_Ridge().save_name_file, Config.Wheat_Ridge().save_name_file, 2, "RUSH_PRIORITY")

                    logger.info("Running the macro for the xloopup rush")
                    RunMacro(Config.Wheat_Ridge().reported_report, Config.macro_path_daily_task, Config.macro_hotkey_wheat_ridge_rush)

                    logger.info(f"Automation Success (Rush). Done automating site {site}. Please checkt the files for validation.")

                elif len(validated_files) == 0:
                    logger.warning(f"No Rush Files found in the folder: {site}")



    except Exception as e:
        print(e)


if __name__ == "__main__":
    Rush()