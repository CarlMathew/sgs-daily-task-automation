import csv
from datetime import datetime, timedelta
import numpy as np
import os
import pandas as pd
import warnings
warnings.filterwarnings('ignore')





class DataPipeline:
    def add_report_type_in_df(self, df:pd.DataFrame, site: str) -> pd.DataFrame:
        """
            This will add all the additional report column in rush priority and late
            This will depends on the site
        """

        if site.title() == "Scott":
            df["GCVOA"] = ''
            df["GCSEMI"] = ''
            df["ORGPREP"] = ''
            df["MSSEMI"] = ''
            df["MSVOA"] = ''
            df["METALS"] = ''
            df["GENCHEM"] = ''
            df["HG"] = ''
        

        ### Additional Features: adding the wheat ridge rush for automation
        if site.title() == "Wheat Ridge":
            pass

        
        return df    
    

    def xlookup_function(self, lookup_value: str, return_array: str, 
                         today_report:pd.DataFrame, 
                         yesterday_report: pd.DataFrame) -> list[str]:
        """Alternative function for the xlookup"""

        data:list[str] = []
        sample_num_ids: list[str] = today_report[lookup_value].to_list()
                
        for sample_num in sample_num_ids:
            yesterday_comment: list = yesterday_report.loc[yesterday_report[lookup_value] == sample_num][return_array].to_list()
        
            if len(yesterday_comment) == 0:
                data.append("NA")
            elif pd.isna(yesterday_comment[0]):
                data.append(" 0 ")
            else:
                try:
                    comment:int = int(yesterday_comment[0])
                    data.append(comment)
                except:
                    data.append(str(yesterday_comment[0]))
        return data
        
    
    # Worklist Pipeline and It's Function

    def count_report_worklist(self, df:pd.DataFrame, excel:csv, report_type: str) -> pd.DataFrame:
        """Count all of the report then store it on csv file. Then sort the due date from oldest to newest"""

        # now: datetime = datetime.now() - timedelta(days=1)
        now: datetime = datetime.now()
        now: datetime = datetime.combine(now.date(), datetime.min.time())
        df["Due Date"] = pd.to_datetime(df["Due Date"], format="%d-%b-%y")
        df = df.sort_values("Due Date")      

        count_all_of_less_than_due_date: int = len(df.loc[df["Due Date"] < now])
        count_all_of_the_data: int = len(df)
        date_today: str = now.date()

        new_data: dict = {
            "Date":date_today,
            "Report_Type" : report_type,
            "Count_Late": count_all_of_less_than_due_date,
            "Total_Count": count_all_of_the_data
        }

        with open(excel, mode="a", newline='') as csvFile:
            writer = csv.DictWriter(csvFile, fieldnames=["Date", "Report_Type", "Count_Late", "Total_Count"])
            writer.writerow(new_data)
        
        df["Due Date"] = df["Due Date"].dt.strftime("%d-%b-%y")

        return df
    
    def remove_lc_from_orgprep(self, df: pd.DataFrame) -> pd.DataFrame:
        """
            This function will remove all of the Products in Orgprep that starts with
            LC, and also include all of its partner that also starts with lc
        """
        splitted_products = list(df["Products"].str.split(", "))

        product_1 = []
        product_2 = []


        for data in splitted_products:
            if len(data) == 1:
                product_1.append(data[0])
                product_2.append(" ")
            else:
                product_1.append(data[0])
                product_2.append(data[1].strip())


        df["Product1"] = product_1
        df["Product2"] = product_2

        df = df.loc[~(df["Product1"].str.contains(r"^LC", regex=True) & (df["Product2"] == " ")) ]
        df = df.loc[~(df["Product1"].str.contains(r"^LC", regex=True) & df["Product2"].str.contains(r"^LC", regex=True))]

        df = df.drop(columns=["Product1", "Product2"])
        
        return df



    def pipeline_for_worklist(self, 
                                files:list[str], 
                                previous_data:str, 
                                qa_samples:str,
                                site:dict,
                                excel_report_path:csv
                            ) -> None:
        """
            Modify and clean the data in worklist extraction
        """
        if not previous_data:
            print("Please download the previous report to continue")
        
        else:
            #Open each file and store it on dataframe
            for file in files:
                # Import All Reports to DataFrame
                df = pd.read_csv(
                    file,
                    sep="\t",
                    engine="python",
                    on_bad_lines="skip"
                )

                # Data Pipeline
                df = df.loc[~df["Acctnum"].isin([qa_samples, "NONE"])] # Remove qa samples
                df = df.drop("Comments", axis = 1) # Drop Comments from every worklist


                #Split the sample num and then insert it into the data
                splitted_data:list[list[str, str]] =list(df["Samplenum"].str.split("-")) 
                split1: list[str] = [data[0] for data in splitted_data]
                split2: list[str] = [data[1] for data in splitted_data]
                df.insert(1, "samplenum", split1)
                df.insert(2, " ", split2)


                # Do the xlookup from the previous report
                report_type: str = site.get(os.path.split(file)[-1].split("_")[0], "")

                # remove the lc from the product column
                if report_type == "ORGPREP":
                    df = self.remove_lc_from_orgprep(df)
                


                yesterday_report = pd.read_excel(previous_data, sheet_name=report_type)
                comments_eta = self.xlookup_function("Samplenum", "Comments/ETA", df, yesterday_report)
                df["Comments/ETA"] = comments_eta


                # Save the count of late to a csv
                df = self.count_report_worklist(df, excel_report_path, report_type)                


                # Additional Cleaning, add more for future reference
                df = df.fillna(" ")
                df["TAT"] = df["TAT"].astype(int)
                df[" "] = df[" "].astype(str)
                df["Comments/ETA"] = df["Comments/ETA"].astype(str)

                df.to_csv(file, sep="\t", index=False)
    
    
    # Rush Pipeline and It's Function


    def pipeline_for_rush(self,
                            files: list[str],
                            previous_data: str,
                            qa_samples: str, 
                            site:str
                         ) -> None: 
        
        """
            Modify and clean the data in rush extraction
        """
        
        for file in files:
            

            df = pd.read_csv(
                    file,
                    sep="\t",
                    engine="python",
                    on_bad_lines="skip"
                )
            

            df = df.loc[df["Account"] != qa_samples]

            df["Comments/ETA"] = ''
            df = self.add_report_type_in_df(df, site)

            last_three_columns_rush = df.tail(3)
            
            df = df.head(len(df) - 3)

            df.loc[df["TAT"].isin(["1", "2", "3", "1*", "2*", "3*"]), "Font"] = "RedAndBold"
            df["Font"].fillna("defualt", inplace=True)


            df["Days Late"] = df["Days Late"].astype(int)
            df["# Spl"] = df["# Spl"].astype(int)

            df.loc[df["Days Late"] >= 1, "FillColor"] = "F4B084"
            df.loc[df["Days Late"] == 0, "FillColor"] = "FFD966"
            df.loc[df["Days Late"] == -1, "FillColor"] = "A9D08E"
            df.loc[df["Days Late"] <=-2, "FillColor"] = "9BC2E6"

            df.loc[df["Job Number"].str.contains(r'X$', regex=True), "FillColor"] = "9BC2E6"


            rush_count: int = len(df.loc[ (df["FillColor"] == "F4B084") & (df["TAT"].isin(["1", "2", "3", "1*", "2*", "3*"]))])
            all_rush_count: int = len(df.loc[(df["TAT"].isin(["1", "2", "3", "1*", "2*", "3*"]))])
            
            priority_count: int = len(df.loc[(df["FillColor"] == "F4B084") & (df["TAT"].isin(["4", "5", "6", "4*", "5*"]))])
            all_priority_count: int = len(df.loc[(df["TAT"].isin(["4", "5", "6", "4*", "5*", "6*"]))])

            print("rush_count: ", rush_count)
            print("all rush count: ", all_rush_count)
            print("priority_count: ", priority_count)
            print("all prio count: ", all_priority_count)

            yesterday_report = pd.read_excel(previous_data, sheet_name = "RUSH_PRIORITY")  

            comments_eta: list[str] = self.xlookup_function("Job Number", "Comments/ETA", df, yesterday_report)
            
            df["Comments/ETA"] = comments_eta
            df["Comments/ETA"] = df["Comments/ETA"].astype(str)

            df.to_csv(file, sep="\t", index=False)
            




            

            




            
            
            



     


        

