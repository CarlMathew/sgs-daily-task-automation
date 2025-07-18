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
    
    def validate_late_file(files:list[str]) -> list[str]:
        """Validate the late file that will be automated"""

        path_list:list[str] = []
        for file in files:
            name_of_file: str = file.split("\\")[-1].split("_")[0]

            if name_of_file.lower() == "late":
                path_list.append(file)
                
        return path_list
    
    def validate_repgen_file(files:list[str]) -> list[str]:
        """Get the data from a folder with a name of repgen"""

        path_list: list[str] = []

        for file in files:
            name_of_file: str = file.split("\\")[-1].split(".")[0]

            if name_of_file.lower() == "repgen":
                path_list.append(file)
        return path_list    
    


if __name__ == "__main__":
    file_management = FileModule()
    files:list[str] = file_management.get_data_from_folder(Config.Wheat_Ridge().raw_file_path)
    validated_files:list[str] =  ValidationOfFiles.validate_repgen_file(files)

    print(validated_files)