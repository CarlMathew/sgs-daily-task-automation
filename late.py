from config import Config
from data_pipeline import DataPipeline
from excel_manipulation import ExcelModule
from file_module import FileModule
from logger_config import setup_late_logger
from report_dictionary import WorklistDictionary
from run_macro import RunMacro
from validation import ValidationOfFiles

logger = setup_late_logger(__name__)

def Late() -> None:
    """Late Automation"""

    try:
        file_management = FileModule()
        pipeline = DataPipeline()
        # worklist = WorklistDictionary()
        excel_module = ExcelModule()
        sites = ['Scott', "Wheat Ridge", "Orlando", 'Dayton']


        for site in sites:

            # -------------------------------------------------Scott Automation----------------------------------------------------------------------------------------------------------------------#
            
            
            # if site == "Scott":
            #     logger.info(f"Automating the late report. Site: {site}")
            #     files: list[str] = file_management.get_data_from_folder(Config.Scott().raw_file_path)

            #     logger.info("Validating Files")
            #     validated_data:list[str] = ValidationOfFiles.validate_late_file(files)

            #     if len(validated_data) == 1:

            #         logger.info("Running Macro Late...")
            #         RunMacro(validated_data, Config.macro_path, Config.macro_hotkey_late)

            #         logger.info("Cleaning the late raw files. (Data Pipeline is running...)")
            #         last_columns = pipeline.pipeline_for_late(validated_data, Config.Scott().yesterday_report_path, Config.Scott().qa_samples, site, Config.Scott().raw_file_path_V2)

            #         logger.info("Inserting the data into the template...")
            #         excel_module.insert_data_from_template(validated_data, Config.Scott().save_name_file, Config.Scott().save_name_file, 3, "LATE", last_columns)
                    
                      # Get the generated csv to insert for repgen
            #         files_for_repgen = file_management.get_data_from_folder(Config.Scott().raw_file_path)
            #         validate_repgen_data: list[str] = ValidationOfFiles.validate_repgen_file(files_for_repgen)

            #         logger.info("Inserting the data into the template (REPGEN)")
            #         excel_module.insert_data_from_template(validate_repgen_data, Config.Scott().save_name_file, Config.Scott().save_name_file, 4, "REPGEN", last_columns)

            #         logger.info("Running the macro for the xloopup late")
            #         RunMacro(Config.Scott().reported_report, Config.macro_path_daily_task, Config.macro_hoteky_scott_late)
                
            #         logger.info(f"Automation Success (Late). Done automating site {site}. Please checkt the files for validation.")

            #     elif len(validated_data) == 0:
            #         logger.warning("No late data found")



            # -------------------------------------------------Wheat Ridge Automation----------------------------------------------------------------------------------------------------------------------#            


            # if site == "Wheat Ridge":
            #     logger.info(f"Automating the late report. Site: {site}")
            #     files: list[str] = file_management.get_data_from_folder(Config.Wheat_Ridge().raw_file_path)

            #     logger.info("Validating Files")
            #     validated_data:list[str] = ValidationOfFiles.validate_late_file(files)


            #     if len(validated_data) == 1:

            #         logger.info("Running Macro Late...")
            #         RunMacro(validated_data, Config.macro_path, Config.macro_hotkey_late)

            #         logger.info("Cleaning the late raw files. (Data Pipeline is running...)")
            #         last_columns = pipeline.pipeline_for_late(validated_data, Config.Wheat_Ridge().yesterday_report_path, Config.Wheat_Ridge().qa_samples, site, Config.Wheat_Ridge().raw_file_path_V2)

            #         logger.info("Inserting the data into the template...")
            #         excel_module.insert_data_from_template(validated_data, Config.Wheat_Ridge().save_name_file, Config.Wheat_Ridge().save_name_file, 3, "LATE", last_columns)
                    
            #         # Get the generated csv to insert for repgen
            #         files_for_repgen = file_management.get_data_from_folder(Config.Wheat_Ridge().raw_file_path)
            #         validate_repgen_data: list[str] = ValidationOfFiles.validate_repgen_file(files_for_repgen)

            #         logger.info("Inserting the data into the template (REPGEN)")
            #         excel_module.insert_data_from_template(validate_repgen_data, Config.Wheat_Ridge().save_name_file, Config.Wheat_Ridge().save_name_file, 4, "REPGEN", last_columns)

            #         logger.info("Running the macro for the xloopup late")
            #         RunMacro(Config.Wheat_Ridge().reported_report, Config.macro_path_daily_task, Config.macro_hoteky_wheat_ridge_late)
            
            #         logger.info(f"Automation Success (Late). Done automating site {site}. Please checkt the files for validation.")

            #     elif len(validated_data) == 0:
            #         logger.warning("No late data found")


            # -------------------------------------------------Orlando Automation----------------------------------------------------------------------------------------------------------------------#

            # if site == "Orlando":
            #     logger.info(f"Automating the late report. Site: {site}")
            #     files: list[str] = file_management.get_data_from_folder(Config.Orlando().raw_file_path)

            #     logger.info("Validating Files")
            #     validated_data:list[str] = ValidationOfFiles.validate_late_file(files)


            #     if len(validated_data) == 1:

            #         logger.info("Running Macro Late...")
            #         RunMacro(validated_data, Config.macro_path, Config.macro_hotkey_late)

            #         logger.info("Cleaning the late raw files. (Data Pipeline is running...)")
            #         last_columns = pipeline.pipeline_for_late(validated_data, Config.Orlando().yesterday_report_path, Config.Orlando().qa_samples, site, Config.Orlando().raw_file_path_V2)

            #         logger.info("Inserting the data into the template...")
            #         excel_module.insert_data_from_template(validated_data, Config.Orlando().save_name_file, Config.Orlando().save_name_file, 3, "LATE", last_columns)
                    
            #         # Get the generated csv to insert for repgen
            #         files_for_repgen = file_management.get_data_from_folder(Config.Orlando().raw_file_path)
            #         validate_repgen_data: list[str] = ValidationOfFiles.validate_repgen_file(files_for_repgen)

            #         logger.info("Inserting the data into the template (REPGEN)")
            #         excel_module.insert_data_from_template(validate_repgen_data, Config.Orlando().save_name_file, Config.Orlando().save_name_file, 4, "REPGEN", last_columns)

            #         logger.info("Running the macro for the xloopup late")
            #         RunMacro(Config.Orlando().reported_report, Config.macro_path_daily_task, Config.macro_hoteky_orlando_late)

            #         logger.info(f"Automation Success (Late). Done automating site {site}. Please checkt the files for validation.")

            #     elif len(validated_data) == 0:
            #         logger.warning("No late data found")

            # -------------------------------------------------Dayton Automation----------------------------------------------------------------------------------------------------------------------#

            if site == "Dayton":
                logger.info(f"Automating the late report. Site: {site}")
                files: list[str] = file_management.get_data_from_folder(Config.Dayton().raw_file_path)

                logger.info("Validating Files")
                validated_data:list[str] = ValidationOfFiles.validate_late_file(files)


                if len(validated_data) == 1:

                    logger.info("Running Macro Late...")
                    RunMacro(validated_data, Config.macro_path, Config.macro_hotkey_late)

                    logger.info("Cleaning the late raw files. (Data Pipeline is running...)")
                    last_column = pipeline.pipeline_for_late(validated_data, Config.Dayton().yesterday_report_path, Config.Dayton().qa_samples, site, Config.Dayton().raw_file_path_V2)

                    logger.info("Inserting the data into the template...")
                    excel_module.insert_data_from_template(validated_data, Config.Dayton().save_name_file, Config.Dayton().save_name_file, 3, "LATE", last_column)
                    
                    # Get the generated csv to insert for repgen
                    files_for_repgen = file_management.get_data_from_folder(Config.Dayton().raw_file_path)
                    validate_repgen_data: list[str] = ValidationOfFiles.validate_repgen_file(files_for_repgen)

                    logger.info("Inserting the data into the template (REPGEN)")
                    excel_module.insert_data_from_template(validate_repgen_data, Config.Dayton().save_name_file, Config.Dayton().save_name_file, 4, "REPGEN", last_column)

                    logger.info("Running the macro for the xloopup late")
                    RunMacro(Config.Dayton().reported_report, Config.macro_path_daily_task, Config.macro_hoteky_dayton_late)

                    logger.info(f"Automation Success (Late). Done automating site {site}. Please checkt the files for validation.")

                elif len(validated_data) == 0:
                    logger.warning("No late data found")

            # ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    except Exception as e:
        logger.critical(f"Error:{e}")



if __name__ == "__main__":
     Late()

