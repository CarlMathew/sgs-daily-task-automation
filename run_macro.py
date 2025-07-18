import os
import win32com.client as win32


def RunMacro(files: list, excel_macro_path: str ,macro_name:str) -> None:
    """
        Execute or run a macro in your excel files
    """
    full_macro_path = os.path.join(os.path.expanduser("~"), "OneDrive - SGS", "Documents", excel_macro_path) # Path of the macro VBA
    try:

        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = True
        
        wb_macro = excel.Workbooks.Open(Filename=full_macro_path) # Open the macro path

        # Iterate all the excel files and activate the macro to edit those excel files
        for file in files:
            wb_target = excel.Workbooks.Open(Filename = file)

            excel.Application.Run(f"'{excel_macro_path}'!{macro_name}")
            
            wb_target.Close(SaveChanges=True)
            

    except KeyboardInterrupt:
        print("Script Interrupted. Cleaning up..")

    except Exception as e:
        print(f"Please check the macro file: {e}")

    finally:
        if wb_macro is not None:
            wb_macro.Close(SaveChanges=False)
        
        if excel is not None:
            excel.Quit()


    
    