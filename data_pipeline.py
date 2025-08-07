import csv
from datetime import datetime, timedelta
import os
import pandas as pd
import re
from utils import Utils
import warnings
warnings.filterwarnings('ignore')





class DataPipeline:

    """Clean the data of the report"""
    
    def xlookup_function(self, lookup_value: str, return_array: str, 
                         today_report:pd.DataFrame, 
                         yesterday_report: pd.DataFrame) -> list[str]:
        """Alternative function for the xlookup"""

        data:list[str] = []
        sample_num_ids: list[str] = today_report[lookup_value].to_list()
                
        for sample_num in sample_num_ids:
            yesterday_comment: list = yesterday_report.loc[yesterday_report[lookup_value] == sample_num][return_array].to_list()

        
            if len(yesterday_comment) == 0:
                data.append("   #N/A   ")
            elif pd.isna(yesterday_comment[0]) or yesterday_comment[0] == 0:
                data.append(" 0 ")
            else:
                try:
                    comment:int = int(yesterday_comment[0])
                    data.append(comment)
                except:
                    data.append(str(yesterday_comment[0]))

        return data
    

        
    
    ######### Worklist Pipeline and It's Function #########



    def count_report_worklist(self, df:pd.DataFrame, excel:csv, report_type: str) -> pd.DataFrame:
        """Count all of the report then store it on csv file. Then sort the due date from oldest to newest"""

        # now: datetime = datetime.now() - timedelta(days=1)

        now: datetime = datetime.now()
        now: datetime = datetime.combine(now.date(), datetime.min.time())

        if report_type == "EXTLCMS":
            df["WorkDate"] = pd.to_datetime(df["WorkDate"], format="%d-%b-%y")
            df = df.sort_values("WorkDate")      
        else:
            df["Due Date"] = pd.to_datetime(df["Due Date"], format="%d-%b-%y")
            df = df.sort_values("Due Date")      

        if report_type == "ORGPREP":
            count_all_of_less_than_due_date: int = len(df.loc[df["Due Date"] <= now])
            count_all_of_the_data: int = len(df)
            date_today: str = now.date()
        elif report_type == "EXTLCMS":
            count_all_of_less_than_due_date: int = len(df.loc[df["WorkDate"] < now])
            count_all_of_the_data: int = len(df)
            date_today: str = now.date()
        else:
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


        if report_type == "EXTLCMS":
            df["WorkDate"] = df["WorkDate"].dt.strftime("%d-%b-%y")
        else:
            df["Due Date"] = df["Due Date"].dt.strftime("%d-%b-%y")

        return df, count_all_of_less_than_due_date, count_all_of_the_data
    
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

    

    def fix_format_comments_genchem(self, df: pd.DataFrame) -> pd.DataFrame:
        """
            Fixed the dateformat of comments (every date string) in genchem into m/d/yy
            eg. (5/11/25, 6/15/25)
        """

        def check_date_and_format(comment:str):
            parsed_dates = pd.to_datetime(comment, errors="coerce")
            if pd.notna(parsed_dates):
                return parsed_dates.strftime("%#m/%#d/%y")
            else:
                return comment
        
        df["Comments/ETA"] = df["Comments/ETA"].apply(check_date_and_format)
        return df 
    


    def filter_all_the_xh_products(self, df: pd.DataFrame, date_logical_statement:bool) -> pd.DataFrame:
        """
            Filter all the products that ends with xh
            if it has multiple values it should also ends with xh
        """


        products_xh: list[dict] = df.loc[ (date_logical_statement) & (df["Products"].str.contains("xh$", case = False)) ] [["Products"]].reset_index().to_dict(orient="records")
        
        filter_products:list = []

        for xh in products_xh:
            index_number: str = xh["index"]
            product_xh: str = xh["Products"]
            split_product_xh:list = product_xh.split(",")
            
            boolean_data:list = []

            for product in split_product_xh:
                if product.strip().endswith(("xh", "XH", "xH", "Xh")):
                    boolean_data.append(True)
                else:
                    boolean_data.append(False)

            data:dict = {}
            if False in boolean_data:
                data["index"] = index_number
                data["status"] = False
            elif False not in boolean_data:
                data["index"] = index_number
                data["status"] = True
            
            filter_products.append(data)


        for product_with_xh in filter_products:

            stat:bool = product_with_xh["status"]
            index_number:str = product_with_xh["index"]
            if stat:
                df.loc[index_number,  "Products_Not_Late"] = 1.0
        

        return df





    def reason_filter_for_dayton(self, df: pd.DataFrame, excel,  report_type:str) -> pd.DataFrame:

        """Add a column on data frame that will filter all the non-late in dayton"""

        now: datetime = datetime.now() # - timedelta(days=1)
        now: datetime = datetime.combine(now.date(), datetime.min.time())


        df["Collected"] = pd.to_datetime(df["Collected"], format="%d-%b-%y")

        df["Due Date"] = pd.to_datetime(df["Due Date"], format="%d-%b-%y")
        df = df.sort_values("Due Date")      


        if report_type == "ORGPREP":
            date_logical_statement = df["Due Date"] <= now
            
        else:
            date_logical_statement = df["Due Date"] < now

        # Count of Late
        count_of_late:int = len(df.loc[(date_logical_statement)])

        # Put gray bg on sample num column (Do Not Count as Late)
            ## If the SampleNum does not starts with JE and JD
        df.loc[
            (date_logical_statement) & 
            ~(df["Samplenum"].str.contains("^JE|^JD", regex=True)), "SampleNum_Not_Late"
        ] = 1

        # Put gray bg on sample num column (Do Not Count as Late)
            ## If the samplenum ends with X
        df.loc[((date_logical_statement) & (df["samplenum"].str.endswith("X")), "SampleNum_Not_Late")]= 1


        # Put gray bg on Acctnum column (Do Not Count as Late)
            ## If Acctnum is equal to ERACOA
        df.loc[((date_logical_statement) & (df["Acctnum"] == "ERACOA" ), "AcctNum_Not_Late")] = 1


        # Put gray bg on Acctnum column (Do Not Count as Late)
            ## If the TAT is 6 and the collected date is less than and equal to 7 days
        if report_type != "ORGPREP":
            filter_date_tat = now - timedelta(days=7)
            df.loc[(date_logical_statement) & (df["TAT"].isin([6, "6", "6*"])) & (df["Collected"] >= filter_date_tat), "TAT_Not_Late"] = 1

        # Put gray bg on Acctnum column (Do Not Count as Late)

        if report_type == "GENCHEM":

            # Put gray on bg in GENCHEM tab if the value of Products is GRAINS (Do not count as late)
            df.loc[(date_logical_statement) & (df["Products"] == "GRAINS"), "Products_Not_Late"] = 1

            filter_date_crany_not_6 = now - timedelta(days=14)
            filter_date_crany_is_6 = now - timedelta(days=Utils.get_total_workday(days=29))

            # Filter all the crany where the TAT is not equal to 6. The days is less than 14
            df.loc[(date_logical_statement) & (df["Acctnum"].str.contains(r"^CRANY" , regex = True)) & ~(df["TAT"].isin([6, "6", "6*"])) & (df["Collected"] >= filter_date_crany_not_6), "AcctNum_Not_Late"] = 1

            # Filter all the crany where the TAT is equal to 6. The days depends on how many workday
            df.loc[(date_logical_statement) & (df["Acctnum"].str.contains(r"^CRANY" , regex = True)) & (df["TAT"].isin([6, "6", "6*"])) & (df["Collected"] >= filter_date_crany_is_6), "AcctNum_Not_Late"] = 1
            
        else:
            filter_date_crany = now - timedelta(days=14)
            df.loc[(date_logical_statement) & (df["Acctnum"].str.contains(r"^CRANY" , regex = True)) & (df["Collected"] >= filter_date_crany), "AcctNum_Not_Late"] = 1

        # Put gray bg on Acctnum column (Do Not Count as Late)
        filter_date_hwin = now - timedelta(days=14)
        df.loc[(date_logical_statement) & (df["Acctnum"].str.contains(r"^HWIN" , regex = True)) & (df["Collected"] >= filter_date_hwin), "AcctNum_Not_Late"] = 1
        

        # Put gray bg on Acctnum column (Do Not Count as Late)
        filter_date_mtx = now - timedelta(days=7)
        df.loc[(date_logical_statement) & (df["Acctnum"].str.contains(r"^MTX" , regex = True)) & (df["Collected"] >= filter_date_mtx), "AcctNum_Not_Late"] = 1
        
        # Put gray bg on Products column if the value ends with xh (Do Not Count as Late)
        # df.loc[(date_logical_statement) & (df["Products"].str.contains(r"xh$", case=False)), "Products_Not_Late"] = 1

        df = self.filter_all_the_xh_products(df, date_logical_statement)

        # Put gray bg on Products if the value starts with B522 (Do Not Count as Late)
        df.loc[(date_logical_statement) & ~(df["Products"].str.contains(",", na=False)) & (df["Products"].str.startswith("B522")), "Products_Not_Late"] = 1

        # Put gray on bg if the collected is last year and greater than 3 months
        df.loc[(date_logical_statement) & (df["Collected"].dt.year == 2024) & (now - df["Collected"] > "100 days"), "Collected_Not_Late"] = 1

    
        if report_type == "ORGPREP":
            # Put gray on bg in ORGPREP tab if the value of TAT is 1,2,3 (Do not count as late)
            df.loc[(date_logical_statement) & (df["TAT"].isin([1,2,3,"1","2","3","1*","2*","3*"])), "TAT_Not_Late"] = 1


        df["Due Date"] = df["Due Date"].dt.strftime("%d-%b-%y")
        df["Collected"] = df["Collected"].dt.strftime("%d-%b-%y")


        # counting all the non late
        count_of_not_actual_late: int = len(
            df.loc[
                (df["SampleNum_Not_Late"] == 1) |
                (df["AcctNum_Not_Late"] == 1) |
                (df["Products_Not_Late"] == 1) |
                (df["TAT_Not_Late"] == 1) | 
                (df["Collected_Not_Late"] == 1)
            ]
        )
        
        # Total Late
        total_late: int = count_of_late - count_of_not_actual_late
        count_all_of_the_data: int = len(df)

        # Inserting data into csv
        new_data: dict = {
            "Date":now.date(),
            "Report_Type" : report_type,
            "Count_Late": total_late,
            "Total_Count": count_all_of_the_data
        }


        with open(excel, mode="a", newline='') as csvFile:
            writer = csv.DictWriter(csvFile, fieldnames=["Date", "Report_Type", "Count_Late", "Total_Count"])
            writer.writerow(new_data)

        
        return df, total_late
    
    def pipeline_for_worklist(self, 
                                files:list[str], 
                                previous_data:str, 
                                qa_samples:str,
                                site:dict,
                                excel_report_path:csv,
                                main_site:str = "Others"
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
                


                # Look up the comments\eta from current to previous
 
                yesterday_report = pd.read_excel(previous_data, sheet_name=report_type)
                comments_eta = self.xlookup_function("Samplenum", "Comments/ETA", df, yesterday_report)
                df["Comments/ETA"] = comments_eta


                # Additional Filters for counting in dayton
                if main_site == "Dayton":
                    df, total_late = self.reason_filter_for_dayton(df, excel_report_path, report_type)

                # Save the count of late to a csv
                else:
                    df, late_count, total_data_count = self.count_report_worklist(df, excel_report_path, report_type)  
                    # worklist_data[""]
                    


                # Added Column in Scott (ORGPREP), Spike
                if main_site == "Scott" and report_type == "ORGPREP":
                    df["Spike added"] = ""
                    spike_added:list[str] = self.xlookup_function("Samplenum", "Spike added", df, yesterday_report)
                    df["Spike added"] = spike_added



                # Format the comments\eta of genchem tab
                if report_type == "GENCHEM":
                    df = self.fix_format_comments_genchem(df)


                
                # Additional Cleaning, add more for future reference
                df = df.fillna(" ")
                df["TAT"] = df["TAT"].astype(int)
                df[" "] = df[" "].astype(str)
                df["Comments/ETA"] = df["Comments/ETA"].astype(str)

                # df.loc[df["Comments/ETA"]=="NA", "Comments/ETA"] = "   #N/A   "

                # save to csv
                df.to_csv(file, sep="\t", index=False)

    
    ######### Rush Pipeline and It's Function #########

    def filter_data_with_tat_6(self, df: pd.DataFrame, tab: str) -> pd.DataFrame:
        """Filter all the data in rush tab and late with tat 6"""

        if tab == "rush":
            account_col: str = "Account"
            fill_color_tab:str = "FillColor"
        elif tab == "late":
            account_col: str = "Account Number"
            fill_color_tab:str = "Fill Color"
        
        data_with_tat_6 = df.loc[ (df[fill_color_tab] == "F4B084") & (df["TAT"].isin([6, "6", "6*"]))][[account_col, "TAT", "Receive Date", fill_color_tab]].reset_index()
        data_with_tat_6 = data_with_tat_6[~ ( data_with_tat_6[account_col].str.contains(r"^CRANY" , regex = True) | data_with_tat_6[account_col].str.contains(r"^HWIN" , regex = True))].to_dict(orient="records")


        for tat in data_with_tat_6:
            index_num = tat["index"]
            received_date = tat["Receive Date"] 
            
            time_difference = (datetime.now() - received_date).days - 7

            if time_difference > 0:
                df.loc[index_num, fill_color_tab] = "F4B084"
            elif time_difference == 0:
                df.loc[index_num, fill_color_tab] = "FFD966"
            elif time_difference == -1:
                df.loc[index_num, fill_color_tab] = "A9D08E"
            elif time_difference <= -2:
                df.loc[index_num, fill_color_tab] = "9BC2E6"


        return df
    
    def filter_special_account(self, df:pd.DataFrame, tab:str, account:str, days:int) -> pd.DataFrame:

        """
            Fillter all the data in rush and late tab based on account tab
        
        """
        now: datetime = datetime.now()
        now: datetime = datetime.combine(now.date(), datetime.min.time())


        if tab == "rush":
            account_col: str = "Account"
            fill_color_tab:str = "FillColor"
        elif tab == "late":
            account_col: str = "Account Number"
            fill_color_tab:str = "Fill Color"


        
        data_with_special_account = df.loc[(df[fill_color_tab] == "F4B084") &(df[account_col].str.contains(rf"^{account}" , regex = True))][[account_col,"Receive Date"]].reset_index().to_dict(orient="records")

        
        for data in data_with_special_account:
            index_num = data["index"]
            receive_date = data["Receive Date"]

            time_difference = (now - receive_date).days - days

            if time_difference > 0:
                df.loc[index_num, fill_color_tab] = "F4B084"

            elif time_difference == 0:
                df.loc[index_num, fill_color_tab] = "FFD966"
            elif time_difference == -1:
                df.loc[index_num, fill_color_tab] = "A9D08E"
            elif time_difference <= -2:
                df.loc[index_num, fill_color_tab] = "9BC2E6"
            

                
        return df


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
        
        if site.title() == "Wheat Ridge":
            df["GCVOA"] = ''
            df["GCSEMI"] = ''
            df["ORGPREP"] = ''
            df["MSSEMI"] = ''
            df["MSVOA"] = ''
            df["METALS"] = ''
            df["GENCHEM"] = ''
            df["HG"] = ''
        
        if site.title() == "Orlando":
            df["GCVOA"] = ''
            df["GCSEMI"] = ''
            df["ORGPREP"] = ''
            df["MSSEMI"] = ''
            df["MSVOA"] = ''
            df["METALS"] = ''
            df["GENCHEM"] = ''
            df["PFAS"] = ''
            df["HG"] = ''
        
        if site.title() == "Dayton":

            df["GCVOA"] = ''
            df["GCSEMI"] = ''
            df["ORGPREP"] = ''
            df["MSSEMI"] = ''
            df["MSVOA"] = ''
            df["MSAIR"] = ''
            df["GCAIR"] = ''
            df["METALS"] = ''
            df["GENCHEM"] = ''
            df["PFAS"] = ''
            df["HG"] = ''
            


        
        return df
    

    def count_report_rush(self, excel:str, rush_count: int, all_rush_count: int, priority_count: int, all_prio_count: int) -> None:
        """Save the data from counting rush_count, total_rush_count, priority_count, total_priority_count"""

        now: datetime = datetime.now()
        date_today: str = now.date()
        new_data: dict = {
            "Date":date_today,
            "Rush Count Open Jobs (1-2-3)" : rush_count,
            "Total Rush Count (1-2-3)": all_rush_count,
            "Priority Count Open Jobs (4-5-6)": priority_count,
            "Total Priority Count (4-5-6)": all_prio_count
        }

        with open(excel, mode="a", newline='') as csvFile:
            writer = csv.DictWriter(csvFile, fieldnames=["Date", "Rush Count Open Jobs (1-2-3)", 
                                                         "Total Rush Count (1-2-3)", "Priority Count Open Jobs (4-5-6)",
                                                         "Total Priority Count (4-5-6)"])
            writer.writerow(new_data)

    
    def remove_pfas_in_rush_orlando(self,df_rush:pd.DataFrame, df_pfas: pd.DataFrame) -> pd.DataFrame:
        """This will remove all the pfas in rush table
            This only apply to Dayton
        """

        df_rush = df_rush.rename(columns={"Job Number": "samplenum"})
        comments_eta = list(map( lambda x: " " if x ==  "   #N/A   " else x ,self.xlookup_function("samplenum", "Comments/ETA", df_rush, df_pfas)))


        df_rush["PFAS_Col"] = comments_eta

        remove_pfas_rush = df_rush.loc[df_rush["TAT"].isin(["7", "7*", "8", "8*", "9", "9*", "10", "10*", 
                                                            "11", "11*", "12", "12*", "13", "13*"])]

        remove_pfas_rush = remove_pfas_rush.loc[df_rush["PFAS_Col"] == " "]

        remove_pfas_rush = remove_pfas_rush.loc[~(df_rush["Incomplete Tests by SGroup"].str.contains(r"\bLCMS\b", regex=True))]
        pfas_index = remove_pfas_rush[["TAT", "Incomplete Tests by SGroup", "PFAS_Col"]].index

        df_rush = df_rush.loc[~(df_rush.index.isin(pfas_index))]

        df_rush = df_rush.rename(columns={"samplenum": "Job Number"}).reset_index().drop("index", axis = 1)
        df_rush = df_rush.drop("PFAS_Col", axis = 1)

        return df_rush


    def remove_pfas_in_rush_dayton(self, df_rush:pd.DataFrame, df_pfas: pd.DataFrame, df_msa: pd.DataFrame, df_gca: pd.DataFrame) -> None:
        """This will remove all the pfas in rush table
            This only apply to Orlando
        """
        

        df_rush = df_rush.rename(columns={"Job Number": "samplenum"})

        

        comments_eta_pfas = list(map( lambda x: " " if x ==  "   #N/A   " else x ,self.xlookup_function("samplenum", "Comments/ETA", df_rush, df_pfas)))
        comments_eta_msa = list(map( lambda x: " " if x ==  "   #N/A   " else x ,self.xlookup_function("samplenum", "Comments/ETA", df_rush, df_msa)))
        comments_eta_gca = list(map( lambda x: " " if x ==  "   #N/A   " else x ,self.xlookup_function("samplenum", "Comments/ETA", df_rush, df_gca)))


        df_rush["PFAS_Col"] = comments_eta_pfas
        df_rush["MSA_Col"] = comments_eta_msa
        df_rush["GCA_Col"] = comments_eta_gca


        remove_late_lab_rush = df_rush.loc[df_rush["TAT"].isin(["7", "7*", "8", "8*", "9", "9*", "10", "10*", 
                                                    "11", "11*", "12", "12*", "13", "13*"])]

        remove_late_lab_rush = remove_late_lab_rush.loc[(df_rush["PFAS_Col"] == " ") | (df_rush["GCA_Col"] == " ") | (df_rush["MSA_Col"] == " ")]

        remove_late_lab_rush = remove_late_lab_rush.loc[
                                                ~(remove_late_lab_rush["Incomplete Tests by SGroup"].str.contains(r"\bLCMS\b", regex=True)) &
                                                ~(remove_late_lab_rush["Incomplete Tests by SGroup"].str.contains(r"\bMSA\b", regex=True)) &
                                                ~(remove_late_lab_rush["Incomplete Tests by SGroup"].str.contains(r"\bGCA\b", regex=True))
                                                ]
        
        
        removing_row_index: list[int] = remove_late_lab_rush[["TAT", "Incomplete Tests by SGroup", "PFAS_Col"]].index

        df_rush = df_rush.loc[~(df_rush.index.isin(removing_row_index))]

        df_rush = df_rush.rename(columns={"samplenum": "Job Number"}).reset_index().drop("index", axis = 1)

        df_rush = df_rush.drop(columns=["PFAS_Col", "MSA_Col", "GCA_Col"])

        return df_rush
    

    def comments_eta_with_date(self, df: pd.DataFrame) -> pd.DataFrame:
        """ 
            Filter all the rows that has due, DD, dd in comments/ETA.
            If date is greater than today
        """
         

        comments_eta_with_due = df.loc[df["FillColor"] == "F4B084"]

        # comments_eta_with_due = comments_eta_with_due.loc[(df["Comments/ETA"].str.contains(r"\bdue\b", case=False)) | (df["Comments/ETA"].str.contains(r"\bdd\b", case=False))][["Comments/ETA"]].reset_index().to_dict(orient="records")
        comments_eta_with_due = comments_eta_with_due.loc[(df["Comments/ETA"].str.contains(r"\bdue\b", case=False)) | (df["Comments/ETA"].str.contains(r"\bdd\b", case=False))]
        comments_eta_with_due = comments_eta_with_due.loc[~(df["Comments/ETA"].str.contains(r"\bpast\b", case=False))][["Comments/ETA"]].reset_index().to_dict(orient="records")

        for comments in comments_eta_with_due:
            comment_eta:str = comments["Comments/ETA"]
            index_num:str = comments["index"]

            dates = re.findall(r'\b\d{1,2}/\d{1,2}\b', comment_eta)


            if dates:
                date = dates[0]
                current_year:datetime = datetime.now().year

                comments_eta_date:str = f"{current_year}/{date}"
                
                comment_date:datetime =  datetime.strptime(comments_eta_date, "%Y/%m/%d")
                
                days_late = (datetime.now() - comment_date).days

                
                if days_late == 0:
                    df.loc[index_num, "FillColor"] = "FFD966"
                
                elif days_late == -1:
                    df.loc[index_num, "FillColor"] = "A9D08E"
                
                elif days_late <= -2:
                    df.loc[index_num, "FillColor"] = "9BC2E6"
        return df


    def additioonal_filter_for_rush(self, df) -> pd.DataFrame:
        """
            Filtering all of the additional request of the client
            
        """
        now: datetime = datetime.now()
        now: datetime = datetime.combine(now.date(), datetime.min.time())

        df["Receive Date"] = pd.to_datetime(df["Receive Date"], format="%d-%b-%y")



        # Filter all the rows with a comments/ETA of DONE or COMPLETED
        df.loc[(df["FillColor"] == "F4B084") & (df["Comments/ETA"].str.contains(r"^Done", case=False)), "FillColor"] = "9BC2E6"
        df.loc[(df["FillColor"] == "F4B084") & (df["Comments/ETA"].str.contains(r"^Completed", case=False)), "FillColor"] = "9BC2E6"
            
        
        df = self.comments_eta_with_date(df)

        # Filter all the non JD and JE
        df.loc[(df["FillColor"] == "F4B084") & ~(df["Job Number"].str.contains("^JE|^JD", regex=True)), "FillColor"] = "9BC2E6"

        # Filter all the rows with ERACOA Account
        df.loc[(df["FillColor"] == "F4B084") & (df["Account"] == "ERACOA" ), "FillColor"] = "9BC2E6"


        #filter_date_tat = now - timedelta(days=7)
        #df.loc[(df["FillColor"] == "F4B084") &  (df["TAT"].isin([6, "6", "6*"])) & (df["Receive Date"] >= filter_date_tat), "FillColor"] = "9BC2E6"

        # filter_date_crany = now - timedelta(days=14)
        # df.loc[(df["FillColor"] == "F4B084") &  (df["Account"].str.contains(r"^CRANY" , regex = True)) & (df["Receive Date"] >= filter_date_crany), "FillColor"] = "9BC2E6"

        # filter_date_hwin = now - timedelta(days=14)
        # df.loc[(df["FillColor"] == "F4B084") &  (df["Account"].str.contains(r"^HWIN" , regex = True)) & (df["Receive Date"] >= filter_date_hwin), "FillColor"] = "9BC2E6"

        # filter_date_mtx = now - timedelta(days=7)
        # df.loc[(df["FillColor"] == "F4B084") & (df["Account"].str.contains(r"^MTX" , regex = True)) & (df["Receive Date"] >= filter_date_mtx), "FillColor"] = "9BC2E6"


        # Filter Account Num that starts with CRANY (14 Days TAT)
        df = self.filter_special_account(df, "rush", "CRANY", 14)
        

        # Filter Account Num that starts with HWIN (14 Days TAT)
        df = self.filter_special_account(df, "rush", "HWIN", 14)

        # Filter Account Num that starts with HWIN (7 Days TAT)
        df = self.filter_special_account(df, "rush", "MTX", 7)

        
        df = self.filter_data_with_tat_6(df, "rush")

        df["Receive Date"] = df["Receive Date"].dt.strftime("%d-%b-%y")

        return df


    def pipeline_for_rush(self,
                            files: list[str],
                            previous_data: str,
                            qa_samples: str, 
                            site:str,
                            save_rush_count:str,
                            save_name_file: str = ""
                         ) -> pd.DataFrame: 
        
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
            
            # Remove all the qa samples in Account and add a comments\eta column
            df = df.loc[df["Account"] != qa_samples]
            df["Comments/ETA"] = ''


            # Look up the comments\eta from current to previous
            yesterday_report = pd.read_excel(previous_data, sheet_name = "RUSH_PRIORITY")  
            comments_eta: list[str] = self.xlookup_function("Job Number", "Comments/ETA", df, yesterday_report)
  
            df["Comments/ETA"] = comments_eta
            df["Comments/ETA"] = df["Comments/ETA"].astype(str)

            df.loc[df["Comments/ETA"]=="NA", "Comments/ETA"] = "   #N/A   "

            # add the columns of per worklist
            df = self.add_report_type_in_df(df, site)
            
            last_three_columns_rush = df.tail(3)
            
            # remove the last three row on the dataframe
            df = df.head(len(df) - 3)

            if site == "Orlando":
                df_pfas = pd.read_excel(save_name_file, sheet_name="LCMSPFAS")
                df = self.remove_pfas_in_rush_orlando(df, df_pfas)
                
                try:
                    df["TAT"] = df["TAT"].astype(int)
                    df["TAT"] = df["TAT"].astype(str)
                except Exception as e:
                    df["TAT"] = df["TAT"].astype(str)
            

            elif site == "Dayton":
                try:
                    df["TAT"] = df["TAT"].astype(int)
                    df["TAT"] = df["TAT"].astype(str)
                except Exception as e:
                    df["TAT"] = df["TAT"].astype(str)
                
                remove_tat_in_rush_dayton: list = [
                    "7", "8", "9", "10", "11", "12",
                    "7*", "8*", "9*", "10*", "11*", "12*"
                ]

                df = df.loc[~(df["TAT"].isin(remove_tat_in_rush_dayton))]

                df_pfas = pd.read_excel(save_name_file, sheet_name="LCMSPFAS")
                df_gca = pd.read_excel(save_name_file, sheet_name="GCAIR")
                df_msa = pd.read_excel(save_name_file, sheet_name="MSAIR")


                df = self.remove_pfas_in_rush_dayton(df, df_pfas, df_msa, df_gca)



            # add a Font with values of "RedAndBold" if in the specified TAT and if not "default" value
            df.loc[df["TAT"].isin(["1", "2", "3", "1*", "2*", "3*"]), "Font"] = "RedAndBold"
            df["Font"].fillna("defualt", inplace=True)


            # change the data type of days late and # Spl
            df["Days Late"] = df["Days Late"].astype(int)
            df["# Spl"] = df["# Spl"].astype(int)

            # add a fill color columns with values based on days late
            df.loc[df["Days Late"] >= 1, "FillColor"] = "F4B084"
            df.loc[df["Days Late"] == 0, "FillColor"] = "FFD966"
            df.loc[df["Days Late"] == -1, "FillColor"] = "A9D08E"
            df.loc[df["Days Late"] <=-2, "FillColor"] = "9BC2E6"


            df.loc[(df["FillColor"] == "F4B084") & (df["Comments/ETA"] == "SUB"), "FillColor"] = "9BC2E6"


            if site == "Dayton":
                df = self.additioonal_filter_for_rush(df)


            
            # change the color of all the subcon to blue. If the end of job number is X
            df.loc[(df["Job Number"].str.contains(r'X$', regex=True)) & (df["FillColor"] == "F4B084"), "FillColor"] = "9BC2E6"

            # Counting
            rush_count: int = len(df.loc[ (df["FillColor"] == "F4B084") & (df["TAT"].isin(["1", "2", "3", "1*", "2*", "3*"]))])
            all_rush_count: int = len(df.loc[(df["TAT"].isin(["1", "2", "3", "1*", "2*", "3*"]))])

            if site == "Orlando":
                priority_count: int = len(df.loc[(df["FillColor"] == "F4B084") & (df["TAT"].isin(["4", "5", "6", "4*", "5*", "13", "13*"]))])
                all_priority_count: int = len(df.loc[(df["TAT"].isin(["4", "5", "6", "4*", "5*", "6*", "13", "13*"]))])
            elif site == "Dayton":
                priority_count: int = len(df.loc[(df["FillColor"] == "F4B084") & (df["TAT"].isin(["4", "4*", "13", "13*" , "6", "6*", "5", "5*"]))])
                all_priority_count: int = len(df.loc[(df["TAT"].isin(["4", "5", "6", "4*", "5*", "6*", "13", "13*"]))])
            else:   
                priority_count: int = len(df.loc[(df["FillColor"] == "F4B084") & (df["TAT"].isin(["4", "5", "6", "4*", "5*"]))])
                all_priority_count: int = len(df.loc[(df["TAT"].isin(["4", "5", "6", "4*", "5*", "6*"]))])

            print("rush_count: ", rush_count)
            print("all rush count: ", all_rush_count)
            print("priority_count: ", priority_count)
            print("all prio count: ", all_priority_count)

            self.count_report_rush(save_rush_count, rush_count, all_rush_count, priority_count, all_priority_count)
            
            #save to csv
            df.to_csv(file, sep="\t", index=False)
            
            # Adding Features
            return last_three_columns_rush

     ######### Late Pipeline and It's Function #########

    def remove_qa_samples_in_late(self, df:pd.DataFrame, qa_samples: str) -> pd.DataFrame:

        """ Remove all the rows with specified qa samples along with its LAB Data: """
        all_lab_data = df.loc[df["Job Number"].str.contains(r"^LAB Data:", regex=True)]
        index_of_labs_data = list(all_lab_data.index)


        for lab_value in index_of_labs_data:
            try:
                for count in range(1, lab_value + 1):
                    job_number = df.iloc[lab_value - count]

                    second_column = list(job_number.index)[1]
                    third_column = list(job_number.index)[2]

                    isSecondColumnExists = job_number[second_column]
                    isThirdColumnExists = job_number[third_column]
                    
                    if pd.notna(isSecondColumnExists) and pd.notna(isThirdColumnExists):
                        df.loc[lab_value, "Parent_JobNumber"] =  job_number["Account Number"]
                        break
            except:
                pass
        df = df.loc[~(df["Account Number"] == qa_samples)]
        df = df.loc[~(df["Parent_JobNumber"] == qa_samples)]

        df = df.drop("Parent_JobNumber", axis = 1)

        return df
    

    def repgen_cleaning_counting(self, df, raw_file_path: str, yesterday_report: pd.DataFrame) -> int:

        """Count the number of repgen and import it into csv"""


        df = df.reset_index()

        df=df.drop("index", axis = 1)

        total_data = len(df)
        all_open_jobs = list(df.loc[df["Days Late"] >= 1].index)

        all_lab_jobs_index = []

        for index_of_open_jobs in all_open_jobs:
            for count in range(1, total_data + 1):
                try:
                    all_job_row = df.iloc[index_of_open_jobs + count]
                    first_column = list(all_job_row.index)[0]
                    starts_with_lab = all_job_row[first_column]

                    if starts_with_lab.strip().startswith("LAB Data:"):
                        all_lab_jobs_index.append(index_of_open_jobs + count)
                        break
                except:
                    pass

        list_of_all_index = list(set(all_lab_jobs_index))

        list_of_all_index.sort()
        for i in list_of_all_index:
            df.loc[i, "Open Job"] = "Yes"


        count_of_repgen = (len(all_open_jobs) + len(df.loc[df["Open Job"] == "Yes"]))   - (len(df.loc[df["Open Job"] == "Yes"]) * 2)

        df = df.drop("Open Job", axis = 1)

        df.loc[df["Job Number"].str.contains(r"---", regex = True, na=False), "Fill Color"] = "FCE4D6"
        repgen_column_name: str = Utils.get_repgen_columns_in_excel(yesterday_report)
        df[repgen_column_name] = ""
        comments_qa:list[str] = self.xlookup_function("Job Number", repgen_column_name, df, yesterday_report)
        df[repgen_column_name] = comments_qa
        #df.loc[(df["Days Late"].isna()) & (df["TAT"].isna()), "Comments (REPGEN)"] = ""
        df = df.rename(columns = {repgen_column_name: "Comments (DT)"})
        

        df.to_csv(raw_file_path + "\\repgen.csv", sep = ",", index = False)
 

        return count_of_repgen
    

    def count_samples_in_late_v1(self, df: pd.DataFrame) -> int:
        # Open Job counting
        open_jobs_count = len(df.loc[df["Fill Color"] == "F4B084"])
        total_open_jobs_count = len(df.loc[df["Fill Color"].notna()])

        print(f"Open Jobs: {open_jobs_count}")
        print(f"Total Open Jobs Count:{total_open_jobs_count}")
        print("-" * 40)
        

        # Late Counting 
        latest = df.loc[df["Fill Color"] == "F4B084"]

        day_today = datetime.today().strftime("%A").lower()
        # day_today = "tuesday"

        if day_today == "monday":
            minimum, mid, maximum = 3,4,5
        elif day_today == "tuesday":
            minimum, mid, maximum = 1,4,5
        elif day_today == "wednesday":
            minimum, mid, maximum = 1,2,5
        else: 
            minimum, mid, maximum = 1,2,3

        count_late_min = len(latest.loc[latest["Days Late"] == minimum])
        count_late_mid = len(latest.loc[latest["Days Late"] == mid])
        count_late_max = len(latest.loc[latest["Days Late"] == maximum]) 

        print(f"Late + 1: {count_late_min}")
        print(f"Late + 2: {count_late_mid}")
        print(f"Late + 3: {count_late_max}")
        print("-" * 40)


        return maximum
    def count_samples_in_late_v2(self, df: pd.DataFrame, maximum) -> None:

        """Count the late in tab and the late > 3"""
        late_greater_3 = len(df.loc[(df["Days Late"] > maximum) & (df["Fill Color"] == "F4B084")])

        late_in_tab = len(df.loc[df["Fill Color"] == "F4B084"])
        
        print(f"Late in Tab: {late_in_tab}")
        print(f"Late > 3: {late_greater_3}")
    

    def additioonal_filter_for_late(self, df) -> pd.DataFrame:
        """
            Filtering all of the additional request of the client (late)
            
        """
        now: datetime = datetime.now()
        now: datetime = datetime.combine(now.date(), datetime.min.time())

        df["Receive Date"] = pd.to_datetime(df["Receive Date"], format="%d-%b-%y")


        
        df.loc[(df["Fill Color"] == "F4B084") & (df["Comments/ETA"] == "DONE"), "Fill Color"] = "9BC2E6"
        df.loc[(df["Fill Color"] == "F4B084") & (df["Comments/ETA"] == "COMPLETED"), "Fill Color"] = "9BC2E6"
        df.loc[(df["Fill Color"] == "F4B084") & (df["Comments/ETA"] == "SUB"), "Fill Color"] = "9BC2E6"

        df.loc[(df["Fill Color"] == "F4B084") & ~(df["Job Number"].str.contains("^JE|^JD", regex=True)), "Fill Color"] = "9BC2E6"
        df.loc[(df["Fill Color"] == "F4B084") & (df["Account Number"] == "ERACOA" ), "Fill Color"] = "9BC2E6"


        # filter_date_tat = now - timedelta(days=7)
        # df.loc[(df["Fill Color"] == "F4B084") &  (df["TAT"].isin([6, "6", "6*"])) & (df["Receive Date"] >= filter_date_tat), "Fill Color"] = "9BC2E6"

        #filter_date_tat = now - timedelta(days=7)
        #df.loc[(df["Fill Color"] == "F4B084") &  (df["TAT"].isin([6, "6", "6*"])) & (df["Receive Date"] >= filter_date_tat), "Fill Color"] = "9BC2E6"


        # filter_date_crany = now - timedelta(days=14)
        # df.loc[(df["Fill Color"] == "F4B084") &  (df["Account Number"].str.contains(r"^CRANY" , regex = True)) & (df["Receive Date"] >= filter_date_crany), "Fill Color"] = "9BC2E6"

        # filter_date_hwin = now - timedelta(days=14)
        # df.loc[(df["Fill Color"] == "F4B084") &  (df["Account Number"].str.contains(r"^HWIN" , regex = True)) & (df["Receive Date"] >= filter_date_hwin), "Fill Color"] = "9BC2E6"

        # filter_date_mtx = now - timedelta(days=7)
        # df.loc[(df["Fill Color"] == "F4B084") & (df["Account Number"].str.contains(r"^MTX" , regex = True)) & (df["Receive Date"] >= filter_date_mtx), "Fill Color"] = "9BC2E6"

        # Filter Account Num that starts with CRANY (14 Days TAT)
        df = self.filter_special_account(df, "late", "CRANY", 14)
        

        # Filter Account Num that starts with HWIN (14 Days TAT)
        df = self.filter_special_account(df, "late", "HWIN", 14)

        # Filter Account Num that starts with HWIN (7 Days TAT)
        df = self.filter_special_account(df, "late", "MTX", 7)

        df = self.filter_data_with_tat_6(df, "late")

        df["Receive Date"] = df["Receive Date"].dt.strftime("%d-%b-%y")

        return df
    

    def xlookup_late_in_rushprio_tab(self, df:pd.DataFrame, current_report:pd.DataFrame) -> pd.DataFrame: 

        """
            Lookup the late tab into rush prio, with a conditions of late only.

            Args:
                df: pd.DataFrame -> current late report
                current_report: pd.DataFrame -> excel file of todays report
            
            Returns:
                pd.DataFrame: return the cleaned data

        """
        current_report = pd.read_excel(current_report, sheet_name="RUSH_PRIORITY")
        df_isLate = df.loc[( df["Days Late"] >= 1) & (df["TAT"].isin([1, "1", "1*", 2, "2", "2*", 3, "3", "3*", 4, "4", "4*", 5, "5", "5*", 6, "6", "6*", 13, "13", "13*"]))]

        index_of_late_lookup = list(df_isLate.index)
        commments_eta_for_late = self.xlookup_function("Job Number", "Comments/ETA", df_isLate, current_report) 


        df.loc[df.index.isin(index_of_late_lookup), "Comments/ETA"] = commments_eta_for_late


        return df

    def pipeline_for_late(self, 
                          files: list[str],
                          previous_report: str,
                          qa_samples:str,
                          site: str,
                          raw_file_path: str,
                          current_report: str = ""
                          ) -> pd.DataFrame:
        
        

        """
            Modify and clean the data in late extraction
        """
        for file in files:
            df = pd.read_csv(
                file,
                sep="\t",
                engine="python"

            )

            # Trimming the dataframe removing the last page of the dataframe
            trim = df.loc[df["Job Number"].str.contains(r"^Total Late", regex=True, na=False)].index[0] - 1
            last_column: pd.DataFrame = df[trim:]
            df = df[:trim]


            # Removing the qa samples
            df = self.remove_qa_samples_in_late(df, qa_samples)
            df = df.drop("Comments", axis = 1)



            # Add a column (Fill color) to fill the background in excel based on days late
            df.loc[df["Days Late"] >= 1, "Fill Color"] = "F4B084"
            df.loc[df["Days Late"] == 0, "Fill Color"] = "FFD966"
            df.loc[df["Days Late"] == -1, "Fill Color"] = "A9D08E"
            df.loc[df["Days Late"] < -1, "Fill Color"] = "9BC2E6"

            # Add a column (Font) to change the font in excel based on tat
            df.loc[df["TAT"].isin(["1", "2", "3", "4", "5", "1*", "2*", "3*", "4*", "5*"]), "Font"] = "RedAndBold"

            if site == "Dayton":
                df.loc[df["TAT"].isin(["6*"]), "Font"] = "PurpleAndBold"
            
            df.loc[df["TAT"].isin(["6"]), "Font"] = "PurpleAndBold"

            df["Font"].fillna("default", inplace=True)
            
            # Count all the late samples (1,2,3) and repgen
            maximum_number:int = self.count_samples_in_late_v1(df)

            #Count and clean Repgen
            yesterday_report_repgen = pd.read_excel(previous_report, sheet_name="REPGEN")
            repgen_count = self.repgen_cleaning_counting(df, raw_file_path, yesterday_report_repgen)
            print(f"Repgen Count: {repgen_count}")
            print("-" * 40)
            
            # Remove DONE in STAT column
            df = df.loc[df['STAT'] != "DONE"]

            # Remove Blanks in DEPT and do not include the lab data in job number
            df = df.loc[~(df["DEPT"].isna() & ~df["Job Number"].str.contains(r"^LAB Data:", regex=True))]
            df = df.loc[~(df["Job Number"].str.contains("---"))]

            # Filter all the sub (Only if the fill color is orange)
            df.loc[(df["DEPT"] == "SC") & (df["Fill Color"] == "F4B084"), "Fill Color"] = "9BC2E6"
            df.loc[(df["STAT"] == "SUB") & (df["Fill Color"] == "F4B084"), "Fill Color"] = "9BC2E6"
            df.loc[ (df["Job Number"].str.contains(r"X$", regex=True)) & (df["Fill Color"] == "F4B084"), "Fill Color"] = "9BC2E6"


            # now: datetime = datetime.now() - timedelta(days=1)
            # date_format_for_tab:str = now.strftime("%y%m%d")
            #yesterday_report = pd.read_excel(previous_report, sheet_name=f"{site}_Late_{date_format_for_tab}")


            # Add the comments/ETA (join today report to yesterday_report)
            df["Comments/ETA"] = ''
            name_of_late_tab: str =  Utils.get_late_tab_name_in_excel(previous_report)
            yesterday_report = pd.read_excel(previous_report, sheet_name=name_of_late_tab)
            comments_eta: list[str] = self.xlookup_function("Job Number", "Comments/ETA", df, yesterday_report)
            df["Comments/ETA"] = comments_eta
            df["Comments/ETA"] = df["Comments/ETA"].astype(str)




            # Dayton additional Filter
            if site == "Dayton":
                df= self.additioonal_filter_for_late(df)
                # Lookup the late tab (Comments/ETA) to rush prio (Dayton For Now)
                df = self.xlookup_late_in_rushprio_tab(df, current_report)


            # Count of late in tab and late > 3
            self.count_samples_in_late_v2(df, maximum_number)

           
            # Add the site in the datframe
            df = self.add_report_type_in_df(df, site)



            # Additional Cleaning
            df.loc[(df["Days Late"].isna()) & (df["TAT"].isna()), "Fill Color"] = None
            df.loc[(df["Days Late"].isna()) & (df["TAT"].isna()), "Font"] = None
            df.loc[(df["Days Late"].isna()) & (df["TAT"].isna()), "Comments/ETA"] = "  "

            # Import the cleaned reports into csv
            df.to_csv(file, sep="\t", index=False)



            return last_column


    ## LANGAN Cleaning
    @staticmethod
    def clean_langan_data(df: pd.DataFrame) -> pd.DataFrame:

        """
            This function will clean the LANGAN Data. 
            Only input the neccessary information on excel
        """
        trim = df.loc[df["Job Number"].str.contains(r"^Total Late", regex=True, na=False)].index[0] - 1

        last_colums = df[trim:]
        df: pd.DataFrame = df[:trim]


        data_langan = df.loc[(df["Account Number"].str.contains(r"^LANGAN", regex=True, case= False)) & (df["Project Description"].str.contains(r"^South Brooklyn", regex=True, case= False))].reset_index().to_dict(orient="records")

        dont_remove:list[int] = []

        for data in data_langan:

            index_num:int = data["index"]   


            if not index_num in dont_remove:
                for index, i in enumerate(range(index_num + 1, len(df) + 1)):

                    try:
                        langan_data = df.loc[df.index == index_num + index ].reset_index().to_dict(orient = "records")
                        row_data:dict = langan_data[0]
                        index_of_data:int = row_data["index"]
                        job_number:str = row_data["Job Number"]
                        account_number_value:str = row_data["Account Number"]
                        project_description_value:str = row_data["Project Description"]


                        if isinstance(account_number_value, str) and account_number_value.lower() == "langan" and project_description_value == "South Brooklyn Marine Terminal, Brooklyn, NY":
                            dont_remove.append(index_of_data)
                        elif job_number.startswith("LAB"):
                            dont_remove.append(index_num)
                            dont_remove.append(index_of_data)
                            break
                        else:
                            break
                                                                                                                                                                                                                    
                    except Exception as e:
                        pass


        df = df.loc[df.index.isin(dont_remove)]

        df = pd.concat([df, last_colums], axis= 0)
        
        # Improve!!!!!!!!!
        df = df.reset_index()
        df = df.drop("index", axis = 1)

        return df
            

            




            
            
            



     


        

