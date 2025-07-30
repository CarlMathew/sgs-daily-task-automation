from file_module import FileModule
from config import Config

class ValidationOfFiles:

    def validate_worklist_file(files:list[str], validation_list:list[str]) -> list[str]:
        """Validate the Worklist file that will be automated."""
        path_list:list[str] = []
        for file in files:
            name_of_file:str = file.split("\\")[-1].split("_")[0]

            if name_of_file in validation_list:
                path_list.append(file)
        return path_list
    

    def validate_rush_file(files:list[str]) -> list[str]:
        """Validate the rush file that will be automated"""

        path_list:list[str] = []
        for file in files:
            name_of_file: str = file.split("\\")[-1].split("-")[0]

            if name_of_file.lower() == "rush":
                path_list.append(file)
                
        return path_list
    
    def validate_late_file(files:list[str], stored_late: int) -> list[str]:
        """Validate the late file that will be automated"""

        path_list_late_only:list[str] = []
        path_list_late_additional:list[str] = []

        ## New Version
        for file in files:
           name_of_file: str = file.split("\\")[-1].split("_")

           if "late" in name_of_file:
               if len(name_of_file) == 2:
                    path_list_late_only.append(file)
               elif len(name_of_file) == 3:
                    path_list_late_additional.append(file)
            

        if stored_late == 1:
            return path_list_late_only
        elif stored_late == 2:
            return path_list_late_additional
        else:
            print("Please only select between 1 and 2")

        ## Old Version:
        # for file in files:
        #     name_of_file: str = file.split("\\")[-1].split("_")[0]

        #     if name_of_file.lower() == "late":
        #         path_list.append(file)
                
        # return path_list

    
    def validate_repgen_file(files:list[str]) -> list[str]:
        """Get the data from a folder with a name of repgen"""

        path_list: list[str] = []

        for file in files:
            name_of_file: str = file.split("\\")[-1].split(".")[0]

            if name_of_file.lower() == "repgen":
                path_list.append(file)
        return path_list    
    

    def validate_tat_file(file:list[str]) -> list[str]:
        """Get all the kpi data from the folder"""
        path_list: list[str] = []

        for file in files:
            name_of_file: str = file.split("\\")[-1].split(".")[0]
            
            if "tat" in name_of_file:
                path_list.append(file)
        return path_list


if __name__ == "__main__":
    file_management = FileModule()
    files:list[str] = file_management.get_data_from_folder(Config.Dayton().raw_file_path)
    validated_files:list[str] =  ValidationOfFiles.validate_tat_file(files)

    print(validated_files)