from sharepoint_access import SharepointModule
from datetime import datetime, timedelta
from config import Config


class DownloadFile(SharepointModule, Config):

    """Download file from sharepoint"""

    def __init__(self, username, password, site_url):
        super().__init__(username, password, site_url)

        self.now:datetime = datetime.now()        
        self.day_today:str = datetime.today().strftime("%A").lower()
        self.date_yesterday = ""
        
        if self.day_today == "monday":
            self.date_yesterday:datetime = datetime.now() - timedelta(days=3)
        else:
            self.date_yesterday:datetime = datetime.now() - timedelta(days=1)
        
        self.year:str = self.now.strftime("%Y")
        self.month:str = self.now.strftime("%m %B")
        self.full_date_yesterday:str = self.date_yesterday.strftime("%y%m%d")

        
    def download_yesterday_report_scott(self, file_url: str, path_url) -> None:
        """download the yesterday late from scott"""

        file_path:str = f"{file_url}/{self.year}/{self.month}/Scott_Late_{self.full_date_yesterday}.xlsx"
        print(file_path)
        local_path:str = f"{path_url}Scott_Late_{self.full_date_yesterday}.xlsx"
        print(local_path)
        self.download_file(file_path, local_path)





        