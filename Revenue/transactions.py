import os
import pandas as pd
import numpy as np
from datetime import datetime
from typing import List
from os.path import isfile
 
class Transaction():
    def __init__(self) -> None:
        self.files = "all.xlsx"
        self.SUBSCRIPTION_STRING = "Subscription Fees"
        self.USAGE_STRING = "Usage fees"
        self.DEVELOPMENT_STRING = "Development fees"
        self.MAINT_STRING = "Recurring maint fees"
        self.CUSTOMISATION_STRING = "Customization"
        self.XERO_DEVELOPEMENT_STRING = "Sales - Development Fee"
        self.XERO_SUBSCRIPTION_STRING = "Sales - Subscription Fees"
        self.XERO_USAGE_STRING = "Sales - Usage Fees"
        self.XERO_CUSTOMISATION_STRING = "Sales - Customization / Adhoc"
        self.XERO_MAINT_STRING = "Sales - Recurring Maintenance Fees"
    
    def get_dataframe(self, filename:str) -> pd.DataFrame:
        df = pd.read_excel(io=filename)
        df = df.drop(df.columns[[i for i in range(1,6)]], axis=1)
        df.columns = ['name', 'debit', 'credit','account']
        return df
    
    def create_record_holder(self) -> dict:
        revenue = dict()
        revenue_type = [self.XERO_DEVELOPEMENT_STRING,
                        self.XERO_SUBSCRIPTION_STRING,
                        self.XERO_USAGE_STRING,
                        self.XERO_CUSTOMISATION_STRING,
                        self.XERO_MAINT_STRING]
        for item in revenue_type:
            revenue[item] = {"credit": 0, "debit":0}
        return revenue
    
    def get_account_type(self,revenue_type:str) -> str:
        if revenue_type == self.XERO_USAGE_STRING:
            return self.USAGE_STRING
        elif revenue_type == self.XERO_SUBSCRIPTION_STRING:
            return  self.SUBSCRIPTION_STRING
        elif revenue_type == self.XERO_DEVELOPEMENT_STRING:
            return  self.DEVELOPMENT_STRING
        elif revenue_type == self.XERO_MAINT_STRING:
            return  self.MAINT_STRING
        elif revenue_type == self.XERO_CUSTOMISATION_STRING:
            return  self.CUSTOMISATION_STRING
        return ""
    
    def get_transaction(self,filename:str) -> List:
        if not isfile(filename):
            print("ERR: File not found")
            return
        transactions:List = []
        df = self.get_dataframe(filename)
        
        for index, row in df.iterrows():
            if pd.isnull(row["name"]):
                record:dict = self.create_record_holder()
                continue
            
            if type(row["name"]) == datetime:
                # entry in the company
                # update the credit and debit for the account type
                record[row['account']]["credit"] += row["credit"]
                record[row['account']]["debit"] += row["debit"]
                continue
            
            entity:str = row["name"]
            if "Total" in entity:
                entity = entity.replace("Total", "").strip()
                if len(entity) <= 0:
                    # reset record entry
                    record:dict = self.create_record_holder()
                    continue
                # Reset all counters and prepare for next entity
                for revenue_type in record:
                    if not self.valid_entry(record[revenue_type]):
                        continue
                    transactions.append([f"\"{entity}\"",
                                         self.get_account_type(revenue_type),
                                         record[revenue_type]["credit"]-record[revenue_type]["debit"],
                                         record[revenue_type]["credit"],
                                         record[revenue_type]["debit"]])
                # reset record entry
                record:dict = self.create_record_holder()
                continue
        return transactions
    def valid_entry(self, entry: dict):
        if entry["credit"] == 0 and entry["debit"] == 0:
            return False
        return True

    def save_as_csv(self,list_item:List) -> None:
        np.savetxt("compiled.csv", 
           list_item,
           delimiter =",", 
           fmt ='%s')
    
    def compile_transactions(self):
        try:
            transactions = self.get_transaction(self.files)
            self.save_as_csv(transactions)
            print("LOG: File has been saved as compiled.csv!!!")
        except Exception as e:
            print(e)
            print("ERR: Failed to compile the file!!!")
 
if __name__ == "__main__":
    t = Transaction()
    t.compile_transactions()
    input('Press ENTER to exit')
            