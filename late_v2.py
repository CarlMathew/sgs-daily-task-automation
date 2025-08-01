from config import Config
from data_pipeline import DataPipeline
from excel_manipulation import ExcelModule
from file_module import FileModule
from logger_config import setup_late_logger
from report_dictionary import WorklistDictionary
from run_macro import RunMacro
from validation import ValidationOfFiles

logger = setup_late_logger(__name__)

def Additional_Late() -> None:
    """Additional Late Automation"""

    try:
        file_management = FileModule()
        # pipeline = DataPipeline()
        worklist = WorklistDictionary()
        excel_module = ExcelModule()
        sites = ["Wheat Ridge", "Orlando", 'Dayton']


        for site in sites:

            # -------------------------------------------------Wheat Ridge Automation----------------------------------------------------------------------------------------------------------------------#            


            if site == "Wheat Ridge":
                logger.info(f"Automating the late report. Site: {site}")
                files: list[str] = file_management.get_data_from_folder(Config.Wheat_Ridge().raw_file_path)

                logger.info("Validating Files")
                validated_data:list[str] = ValidationOfFiles.validate_late_file(files, 2)


                if len(validated_data) > 0:


                    logger.info("Running Macro Late...")
                    RunMacro(validated_data, Config.macro_path, Config.macro_hotkey_late)

                    # logger.info("Cleaning the late raw files. (Data Pipeline is running...)")
                    # last_column = pipeline.pipeline_for_late(validated_data, Config.Dayton().yesterday_report_path, Config.Dayton().qa_samples, site, Config.Dayton().raw_file_path_V2)

                    logger.info("Inserting the data into the template...")
                    excel_module.insert_data_from_template(validated_data, Config.Wheat_Ridge().save_name_file, Config.Wheat_Ridge().save_name_file, 6, worklist.Wheat_Ridge_Late)
                    
                    # # Get the generated csv to insert for repgen
                    # files_for_repgen = file_management.get_data_from_folder(Config.Dayton().raw_file_path)
                    # validate_repgen_data: list[str] = ValidationOfFiles.validate_repgen_file(files_for_repgen)

                    # logger.info("Inserting the data into the template (REPGEN)")
                    # excel_module.insert_data_from_template(validate_repgen_data, Config.Dayton().save_name_file, Config.Dayton().save_name_file, 4, "REPGEN", last_column)

                    logger.info("Running the macro for the xloopup additional late")
                    RunMacro(Config.Wheat_Ridge().reported_report, Config.macro_path_daily_task, Config.macro_additional_late_wheat_ridge)

                    logger.info(f"Automation Success (Late V2). Done automating site {site}. Please checkt the files for validation.")

                elif len(validated_data) == 0:
                    logger.warning("No late data found")


            # -------------------------------------------------Orlando Automation----------------------------------------------------------------------------------------------------------------------#

            # if site == "Orlando":
            #     logger.info(f"Automating the additional late report. Site: {site}")
            #     files: list[str] = file_management.get_data_from_folder(Config.Orlando().raw_file_path)

            #     logger.info("Validating Files")
            #     validated_data:list[str] = ValidationOfFiles.validate_late_file(files, 2)

            #     if len(validated_data) > 0:


            #         logger.info("Running Macro Late...")
            #         RunMacro(validated_data, Config.macro_path, Config.macro_hotkey_late)

            #         # logger.info("Cleaning the late raw files. (Data Pipeline is running...)")
            #         # last_column = pipeline.pipeline_for_late(validated_data, Config.Dayton().yesterday_report_path, Config.Dayton().qa_samples, site, Config.Dayton().raw_file_path_V2)

            #         logger.info("Inserting the data into the template...")
            #         excel_module.insert_data_from_template(validated_data, Config.Orlando().save_name_file, Config.Orlando().save_name_file, 6, worklist.Orlando_Late)
                    
            #         # # Get the generated csv to insert for repgen
            #         # files_for_repgen = file_management.get_data_from_folder(Config.Dayton().raw_file_path)
            #         # validate_repgen_data: list[str] = ValidationOfFiles.validate_repgen_file(files_for_repgen)

            #         # logger.info("Inserting the data into the template (REPGEN)")
            #         # excel_module.insert_data_from_template(validate_repgen_data, Config.Dayton().save_name_file, Config.Dayton().save_name_file, 4, "REPGEN", last_column)

            #         logger.info("Running the macro for the xloopup additional late")
            #         RunMacro(Config.Orlando().reported_report, Config.macro_path_daily_task, Config.macro_additional_late_orlando)

            #         logger.info(f"Automation Success (Late V2). Done automating site {site}. Please checkt the files for validation.")

            #     elif len(validated_data) == 0:
            #         logger.warning("No late data found")

            # -------------------------------------------------Dayton Automation----------------------------------------------------------------------------------------------------------------------#

            if site == "Dayton":
                logger.info(f"Automating the additional late report. Site: {site}")
                files: list[str] = file_management.get_data_from_folder(Config.Dayton().raw_file_path)

                logger.info("Validating Files")
                validated_data:list[str] = ValidationOfFiles.validate_late_file(files, 2)

                if len(validated_data) > 0:


                    logger.info("Running Macro Late...")
                    RunMacro(validated_data, Config.macro_path, Config.macro_hotkey_late)

                    # logger.info("Cleaning the late raw files. (Data Pipeline is running...)")
                    # last_column = pipeline.pipeline_for_late(validated_data, Config.Dayton().yesterday_report_path, Config.Dayton().qa_samples, site, Config.Dayton().raw_file_path_V2)

                    logger.info("Inserting the data into the template...")
                    excel_module.insert_data_from_template(validated_data, Config.Dayton().save_name_file, Config.Dayton().save_name_file, 6, worklist.Dayton_Late)
                    
                    # # Get the generated csv to insert for repgen
                    # files_for_repgen = file_management.get_data_from_folder(Config.Dayton().raw_file_path)
                    # validate_repgen_data: list[str] = ValidationOfFiles.validate_repgen_file(files_for_repgen)

                    # logger.info("Inserting the data into the template (REPGEN)")
                    # excel_module.insert_data_from_template(validate_repgen_data, Config.Dayton().save_name_file, Config.Dayton().save_name_file, 4, "REPGEN", last_column)

                    logger.info("Running the macro for the xloopup additional late")
                    RunMacro(Config.Dayton().reported_report, Config.macro_path_daily_task, Config.macro_additional_late_dayton)

                    logger.info(f"Automation Success (Late V2). Done automating site {site}. Please checkt the files for validation.")

                elif len(validated_data) == 0:
                    logger.warning("No late data found")

            # ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    except Exception as e:
        logger.critical(f"Error:{e}")



if __name__ == "__main__":
     Additional_Late()

