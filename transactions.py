import os
import pandas as pd
import numpy as np
from datetime import datetime
from typing import List
from os.path import isfile
    
class Menu():
    def __init__(self) -> None:
        pass
    def change_name(filein, filename, newfilename):
        try:
            os.rename(filein, filein.replace(filename,newfilename))
            print ("--- File Renamed ---")
        except Exception as e:
            print(e)
            print("--- ERROR ---")
    def change_name_menu() -> None:
        while (True):
            print("""
                1. usage.xlsx
                2. subscription.xlsx
                3. development.xlsx
                4. maint.xlsx
                5. customization.xlsx
                6. return main menu""")
            userin = input("Which rename option: ")
            if userin == "6":
                return
            
            filein:str = input("Drag and drop file").replace("\"", "").replace("\'", "")
            print ("--- File Received ---")
            filename:str = os.path.basename(os.path.normpath(filein))
            filename = filename.strip().replace("\"", "")
            
            if userin == "1":
                Menu.change_name(filein, filename, "usage.xlsx")
            elif userin == "2":
                Menu.change_name(filein, filename, "subscription.xlsx")
            elif userin == "3":
                Menu.change_name(filein, filename, "development.xlsx")
            elif userin == "4":
                Menu.change_name(filein, filename, "maint.xlsx")
            elif userin == "5":
                Menu.change_name(filein, filename, "customization.xlsx")
            
    def main_menu() -> None:
        while True:
            print("""
                1. Change name
                2. Compile
                3. Exit""")
            userin = input("Which rename option: ")
            if userin == "1":
                Menu.change_name_menu()
            elif userin == "2":
                t = Transaction()
                t. compile_transactions()
            else:
                return
    def start() -> None:
        Menu.main_menu()

class Transaction():
    def __init__(self) -> None:
        self.files = ["usage.xlsx","subscription.xlsx","development.xlsx","maint.xlsx","customization.xlsx"]
        self.SUBSCRIPTION_STRING = "Subscription fees"
        self.USAGE_STRING = "Usage fees"
        self.DEVELOPMENT_STRING = "Development fees"
        self.MAINT_STRING = "Recurring maint fees"
        self.CUSTOMISATION_STRING = "Customization"
    
    def get_dataframe(self, filename:str) -> pd.DataFrame:
        df = pd.read_excel(io=filename)
        df = df.drop(df.columns[[i for i in range(1,6)] + [8]], axis=1)
        df.columns = ['name', 'debit', 'credit']
        return df
    
    def get_transaction(self,filename:str) -> List:
        if not isfile(filename):
            return
        transactions:List = []
        entity:str = ""
        credit:float = 0
        debit:float = 0
        account_type:str = self.get_account_type(filename)
        df = self.get_dataframe(filename)
        
        for index, row in df.iterrows():
            if pd.isnull(row[0]):
                credit:float = 0
                debit:float = 0
                continue
            
            if type(row[0]) == datetime:
                credit += row["credit"]
                debit += row["debit"]
                continue
            
            string:str = row[0]
            if "Total" in string:
                string = string.replace("Total", "").strip()
                # Reset all counters and prepare for next entity
                if len(string) > 0:
                    transactions.append([f"\"{string}\"", account_type, credit-debit, credit, debit])
                credit:float = 0
                debit:float = 0
                continue
        return transactions

    def get_account_type(self,filename:str) -> str:
        if filename == "usage.xlsx":
            return self.USAGE_STRING
        elif filename == "subscription.xlsx":
            return  self.SUBSCRIPTION_STRING
        elif filename == "development.xlsx":
            return  self.DEVELOPMENT_STRING
        elif filename == "maint.xlsx":
            return  self.MAINT_STRING
        elif filename == "customization.xlsx":
            return  self.CUSTOMISATION_STRING
        return ""        

    def get_transactions(self) -> List:
        transactions:List = []
        for file in self.files:
            # store transactions
            try:

                transactions.extend(self.get_transaction(file))
                print(f"{file}: Processing completed")
            except:
                print(f"{file}: Does not exist in current folder")
        return transactions

    def save_as_csv(self,list_item:List) -> None:
        """Uses numpy to save to csv

        Args:
            list_item (List): Lists of transaction entry to save
        """
        np.savetxt("compiled.csv", 
           list_item,
           delimiter =",", 
           fmt ='%s')
    
    def compile_transactions(self) -> None:
        """Gets and saves all transaction type
        """
        transactions = self.get_transactions()
        self.save_as_csv(transactions)

if __name__ == "__main__":
    Menu.start()
    input('Press ENTER to exit')
            