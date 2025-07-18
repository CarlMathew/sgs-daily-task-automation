
import datetime
import json
import os
import pandas as pd
import shutil


# File Management module. This can be used to automate your files. Removing, moving, getting infos.
# You can add for additional Features.

class FileModule:
    def __init__(self):
        self.root_path:str = os.path.dirname(os.path.abspath(__file__))
    
    def get_data_from_folder(self, folder_path: str) -> list[str]:

        """Return a list with names of files inside the folder path"""

        folder: str = self.root_path + folder_path

        files:list[str] = [self.root_path + folder_path + "\\" + file_name for file_name in os.listdir(folder)] # Store all the file path into array
        

        # if folder have files
        if len(files) > 0:
            return files
        # if folder doesnt have any files inside
        else:
            return []
    

    def get_date_created(self, folder_path:str) -> list[datetime]:

        """Return a list with time of date creation of files"""

        folder: str = self.root_path + folder_path
        files: list = os.listdir(folder)

        if len(files)>0:
            files_creation_date: list[datetime] = [
            datetime.datetime.fromtimestamp(
                os.path.getctime(folder + "//" + file)
            ).strftime("%Y-%m-%d")
            for file in files
        ]
            return files_creation_date
        else: 
            return []

    
    def move_file(self, files_from: str, destination:str, isEndWith:bool=False, endswith:str ="") -> list[str]: 
        """This function will move all of your files inside the folder into another folder"""
        files: list[str] = self.get_data_from_folder(files_from)
        remaining_files: list[str] = []
        for file in files:
            _, file_name = os.path.split(file)
            desination_file_name: str = self.root_path + "\\" + destination + "\\" + file_name
            if endswith and file_name.endswith(endswith):
                os.rename(file, desination_file_name)
            else: 
                os.rename(file, desination_file_name)
        remaining_files.append(desination_file_name)

        return self.get_data_from_folder(files_from)

    def is_created(self, folder_path: str = "") -> list[bool]:
        """
            Return a list of boolean if the date is equal to the date creation of files
        """

        date: str = datetime.datetime.now().strftime("%Y-%m-%d")
        files_creation_date: list[datetime] = self.get_date_created(folder_path=folder_path)

        if len(files_creation_date)> 0:

    
            is_createToday_list: list[bool] = [date == files for files in files_creation_date]
            
            return is_createToday_list

    
    def files_info(self, folder_path:str = "") -> list[dict]:
        """Return a list dictionary about the information of the files"""

        name_of_files:list[str] = self.get_data_from_folder(folder_path)

        if len(name_of_files) > 0:

            date_created:list[str] = self.get_date_created(folder_path=folder_path)
            isCreated: list[bool] = self.is_created(folder_path)


            dataFrame: pd = pd.DataFrame({
                "File Name": name_of_files,
                "Created Date": date_created,
                "Is Created": isCreated
            })
                    
            return json.loads(dataFrame.to_json(orient="records"))
        
        
    



