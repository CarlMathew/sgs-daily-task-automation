from config import Config
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
        sites = ['Scott', "Wheat Ridge", "Orlando", "Dayton"]


        for site in sites:
            # ---------------------------------------------------------------------------SCOTT AUTOMATION----------------------------------------------------------------------------------------#
            # if site == "Scott":
            #     logger.info(f"Automating the rush report. Site: {site}")
            #     files:list[str] = file_management.get_data_from_folder(Config.Scott().raw_file_path)

                
            #     logger.info(f"Validating Files")
            #     validated_files:list[str] =  ValidationOfFiles.validate_rush_file(files)

            #     if len(validated_files) == 1:
                    
            #         logger.info("Running Macro Rush...")
            #         RunMacro(validated_files, Config.macro_path, Config.macro_hotkey_rush)

            #         logger.info("Cleaning the rush raw files. (Data Pipeline is running...)")
            #         last_three_columns = pipeline.pipeline_for_rush(validated_files, Config.Scott().yesterday_report_path, Config.Scott().qa_samples, site, Config.Scott().rush_counting_data)


            #         logger.info("Inserting the data into the template...")
            #         excel_module.insert_data_from_template(validated_files, Config.Scott().save_name_file, Config.Scott().save_name_file, 2, "RUSH_PRIORITY",last_three_columns)
                                        
            #         logger.info("Running the macro for the xloopup rush")
            #         RunMacro(Config.Scott().reported_report, Config.macro_path_daily_task, Config.macro_hotkey_scott_rush)
                    

            #         logger.info(f"Automation Success (Rush). Done automating site {site}. Please checkt the files for validation.")


            #     elif len(validated_files) == 0:
            #         logger.warning(f"No Rush Files found in the folder: {site}")

            # ---------------------------------------------------------------------------WHEAT RIDGE AUTOMATION----------------------------------------------------------------------------------------#

            # if site == "Wheat Ridge":
            #     logger.info(f"Automating the rush report. Site: {site}")
            #     files:list[str] = file_management.get_data_from_folder(Config.Wheat_Ridge().raw_file_path)



            #     logger.info(f"Validating Files")
            #     validated_files:list[str] =  ValidationOfFiles.validate_rush_file(files)


            #     if len(validated_files) == 1:

            #         logger.info("Running Macro Rush...")
            #         RunMacro(validated_files, Config.macro_path, Config.macro_hotkey_rush)

            #         logger.info("Cleaning the rush raw files. (Data Pipeline is running...)")
            #         last_three_columns = pipeline.pipeline_for_rush(validated_files, Config.Wheat_Ridge().yesterday_report_path, Config.Wheat_Ridge().qa_samples, site, Config.Wheat_Ridge().rush_counting_data)


            #         logger.info("Inserting the data into the template...")
            #         excel_module.insert_data_from_template(validated_files, Config.Wheat_Ridge().save_name_file, Config.Wheat_Ridge().save_name_file, 2, "RUSH_PRIORITY", last_three_columns)

            #         logger.info("Running the macro for the xloopup rush")
            #         RunMacro(Config.Wheat_Ridge().reported_report, Config.macro_path_daily_task, Config.macro_hotkey_wheat_ridge_rush)

            #         logger.info(f"Automation Success (Rush). Done automating site {site}. Please checkt the files for validation.")

            #     elif len(validated_files) == 0:
            #         logger.warning(f"No Rush Files found in the folder: {site}")


            # ---------------------------------------------------------------------------ORLANDO AUTOMATION----------------------------------------------------------------------------------------#

            # if site == "Orlando":
            #     logger.info(f"Automating the rush report. Site: {site}")
            #     files:list[str] = file_management.get_data_from_folder(Config.Orlando().raw_file_path)

            #     logger.info(f"Validating Files")
            #     validated_files:list[str] =  ValidationOfFiles.validate_rush_file(files)


            #     if len(validated_files) == 1:

            #         logger.info("Running Macro Rush...")
            #         RunMacro(validated_files, Config.macro_path, Config.macro_hotkey_rush)

            #         logger.info("Cleaning the rush raw files. (Data Pipeline is running...)")
            #         last_three_columns = pipeline.pipeline_for_rush(validated_files, Config.Orlando().yesterday_report_path, Config.Orlando().qa_samples, site, Config.Orlando().rush_counting_data, Config.Orlando().save_name_file)

            #         logger.info("Inserting the data into the template...")
            #         excel_module.insert_data_from_template(validated_files, Config.Orlando().save_name_file, Config.Orlando().save_name_file, 2, "RUSH_PRIORITY", last_three_columns)

            #         logger.info("Running the macro for the xloopup rush")
            #         RunMacro(Config.Orlando().reported_report, Config.macro_path_daily_task, Config.macro_hotkey_orlando_rush)

            #         logger.info(f"Automation Success (Rush). Done automating site {site}. Please checkt the files for validation.")

            #     elif len(validated_files) == 0:
            #         logger.warning(f"No Rush Files found in the folder: {site}")

            
            # ---------------------------------------------------------------------------DAYTON AUTOMATION----------------------------------------------------------------------------------------#
                
            if site == "Dayton":
                logger.info(f"Automating the rush report. Site: {site}")
                files:list[str] = file_management.get_data_from_folder(Config.Dayton().raw_file_path)

                logger.info(f"Validating Files")
                validated_files:list[str] =  ValidationOfFiles.validate_rush_file(files)


                if len(validated_files) == 1:

                    logger.info("Running Macro Rush...")
                    RunMacro(validated_files, Config.macro_path, Config.macro_hotkey_rush)

                    logger.info("Cleaning the rush raw files. (Data Pipeline is running...)")
                    last_three_columns = pipeline.pipeline_for_rush(validated_files, Config.Dayton().yesterday_report_path, Config.Dayton().qa_samples, site, Config.Dayton().rush_counting_data, Config.Dayton().save_name_file)

                    logger.info("Inserting the data into the template...")
                    excel_module.insert_data_from_template(validated_files, Config.Dayton().save_name_file, Config.Dayton().save_name_file, 2, "RUSH_PRIORITY", last_three_columns)

                    logger.info("Running the macro for the xloopup rush")
                    RunMacro(Config.Dayton().reported_report, Config.macro_path_daily_task, Config.macro_hotkey_dayton_rush)

                    logger.info(f"Automation Success (Rush). Done automating site {site}. Please checkt the files for validation.")

                elif len(validated_files) == 0:
                    logger.warning(f"No Rush Files found in the folder: {site}")

            # -------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    except Exception as e:
        logger.critical(f"Error: {e}")


if __name__ == "__main__":
    Rush()