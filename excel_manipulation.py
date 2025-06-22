from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, numbers, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
import pandas as pd




class ExcelModule:
    
    def __init__(self):
        self.worklist_row_height: int = 21
        self.rush_row_height: int = 24
        

    def process_automation_worklist(self, ws, columns:list[str], data:list[tuple]) -> bool:
        """
            Start the automation of copy values from excel to excel on worklist data
        """
        try:
            for row in range(1, len(data) + 1):
                ws.row_dimensions[row].height = self.worklist_row_height
                for col in range(len(columns)):
                    char: str = get_column_letter(col + 1)
                    
                    cell: str = char + str(row)
                    value:str = data[row -1][col]
                    ws[cell].alignment = Alignment(horizontal="center", vertical="center")
                    

                    if char == "C":
                        try:
                            split_1_value: str | int = int(value)
                            ws[cell] = split_1_value
                            ws[cell].number_format = numbers.FORMAT_TEXT
                        except:
                            ws[cell] = str(value)
                    elif char == "H":
                        ws[cell] = value
                        ws[cell].number_format = "0"
                    elif char == "K":
                        # date_value: datetime = datetime.strptime(value, "%Y-%m-%d")
                        ws[cell] = value
                        ws[cell].number_format = numbers.FORMAT_DATE_DDMMYY
                    elif char == "N":
                        try:
                            comments_eta_value: str | int = int(value)
                            ws[cell] = comments_eta_value
                            ws[cell].number_format = numbers.FORMAT_TEXT
                        except:
                            ws[cell] = str(value)

                    else:
                        try:
                            value: str | int = int(value)
                            ws[cell] = value
                            ws[cell].number_format = numbers.FORMAT_TEXT
                        except:
                            ws[cell] = str(value)

            return True
        except Exception as e:
            return False


    def process_automation_rush(self, ws, df) -> bool:
        """Start the automation of copy values from excel to excel on rush data"""
        font_and_cell_color_df: pd.DataFrame = df[["Font", "FillColor"]]
        remaining_columns_df: pd.DataFrame = df.drop(columns = ["Font", "FillColor"])

       
        remaining_columns_df["Comments/ETA"] = remaining_columns_df["Comments/ETA"].astype(str)

        font_and_cell_color_dict: list[dict] = font_and_cell_color_df.to_dict(orient="records")
        remaining_columns_dict: list[dict] = remaining_columns_df.to_dict(orient = "records")
        
        self.change_style_rush(ws, df)
        
        columns: list[str] = list(remaining_columns_df.columns)
        rows_data: list[tuple] = [tuple(rows.values()) for rows in remaining_columns_dict]

        rows_data.insert(0, columns)
        
        for data, styles in zip(rows_data, font_and_cell_color_dict):
            for row in range(1, len(rows_data) + 1):
                for col in range(len(columns)):
                    char: str = get_column_letter(col + 1)
                    cell_name: str = char + str(row)
                    cell_value:str = rows_data[row - 1][col]

                    
                    ws[cell_name].alignment = Alignment(horizontal="center", vertical="center")

                    if char == "G":
                        try: 
                            number_value: int = round(float(cell_value))
                            ws[cell_name] = number_value
                        except:
                            string_value: str = cell_value
                            ws[cell_name] = string_value
                    elif char == "H" or char == "I":
                        try:
                            number_value: int = int(cell_value)
                            ws[cell_name] = number_value
                        except:
                            string_value: str = str(cell_value)
                            ws[cell_name] = string_value
                    elif char == "L": 
                        ws[cell_name].alignment = Alignment(horizontal="center", vertical="center")
                        string_value: str = cell_value
                        ws[cell_name] = string_value
                    elif char == "M":
                        ws[cell_name].alignment = Alignment(horizontal="center", vertical="center")
                        try:
                            number_value: int = round(float(cell_value))
                            ws[cell_name] = number_value
                        except:
                            if cell_value == "nan":
                                no_value: str = " "
                                ws[cell_name] = no_value
                            else:
                                ws[cell_name] = cell_value
                    else:
                        ws[cell_name] = cell_value

        max_column:int = ws.max_column
        last_col_letter:str = ws.cell(row=1, column=max_column).column_letter
        table_rush_range:str = f"A1:{last_col_letter}{len(rows_data)}"

        table_ref = Table(displayName="Rush_Table", ref=table_rush_range)

        style_table = TableStyleInfo(name="TableStyleLight1")

        table_ref.tableStyleInfo = style_table
        
        ws.add_table(table_ref)
        
    

    def change_style_rush(self, ws, df) -> None:
        """ Update the font and fill color of each cell """
        dataframe_dict: list[dict] = df.to_dict(orient = "records")
        length_of_columns: int = len(df.columns) - 2
        length_of_rows: int = len(dataframe_dict)


        for row in range(2, length_of_rows + 2):
            ws.row_dimensions[row].height = self.rush_row_height
            fonts: str = dataframe_dict[row - 2]["Font"]
            fill_color:str = dataframe_dict[row-2]["FillColor"]
            

            for col in range(length_of_columns):
                char: str = get_column_letter(col + 1)
                if fonts == "RedAndBold":
                    red_bold_font = Font(color='FF0000', bold=True)
                    
                    ws[char + str(row)].font = red_bold_font
                    

                fill_bg = PatternFill(fill_type="solid", fgColor=fill_color)
                ws[char + str(row)].fill = fill_bg




    def insert_data_from_template(self, 
                                  files: list[str],
                                  from_excel:str="",
                                  to_excel:str="",
                                  automation_type:int=1,
                                  site:dict={}) -> None :
        """
            Transfer data from excel file to excel file
        """
        wb = load_workbook(from_excel)

        for file in files:
            if automation_type == 1:
                report_type: str = site.get(os.path.split(file)[-1].split("_")[0], "")
                df = pd.read_csv(
                    file,
                    sep="\t",
                    engine="python"

                )
                ws = wb[report_type]
                
                max_column = ws.max_column
                last_col_letter = ws.cell(row=1, column=max_column).column_letter
                ws.auto_filter.ref = f"A1:{last_col_letter}1"

                df_to_dict: list[dict] = df.to_dict(orient="records")
                columns: list[str] = tuple(df_to_dict[0].keys())
                data: list[tuple] = [tuple(data.values()) for data in df_to_dict]
                data.insert(0, columns)
                isFinished:bool = self.process_automation_worklist(ws, columns=columns, data=data)

                wb.save(to_excel)


            elif automation_type == 2:
                report_type = "RUSH_PRIORITY"
                df = pd.read_csv(
                    file,
                    sep="\t",
                    engine="python"

                )
                ws = wb[report_type]

                is_finished = self.process_automation_rush(ws, df)
                
                wb.save(to_excel)



        wb.close()




