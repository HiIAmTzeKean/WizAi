import pandas as pd
from typing import Tuple, Dict, List
import numpy as np

class Profit_and_Loss():
    def __init__(self) -> None:
        self.client_file: str = "client.xlsx"
        self.profit_loss_file: str = "all.xlsx"
        self.CHURNED = "CHURNED"
        self.create_client_mapping()
        self.pl_df = self.get_profit_loss_df()
    
    def restore_orignal_columns(self) -> None:
        self.pl_df["Description"] = self.df_original["Description"].values
        self.pl_df["Contact"] = self.df_original["Contact"].values
    
    def clean_string(string:str) -> str:
        return string.lower().replace(",","").replace(".","")
    
    def clean_customer_code(string:str) -> str:
        return string.lstrip("0")
    
    def get_profit_loss_df(self) -> pd.DataFrame:
        df = pd.read_excel(io=self.profit_loss_file)

        # Update header
        df.columns = df.iloc[3]
        holding_entity = df.iloc[0,0]

        # Drop first 3 rows as it contains file meta
        df = df.drop(index=[0,1,2,3])
        df.dropna()
        df["Client name"] = ""
        df["Client code"] = ""
        df["To Check"] = ""
        
        # Preserve old data
        self.df_original =  df[['Description', 'Contact']].copy()
        df['Description'] = df['Description'].str.lower()
        df['Contact'] = df['Contact'].str.lower()
        df = df.assign(Entity=holding_entity)
        return df
    
    def create_client_mapping(self) -> None:
        """Creates variables as helper for mapping
        """
        # Read excel file
        df = pd.read_excel(io=self.client_file,header=0)
        # Clean file
        df["Client Code"] = pd.to_numeric(df["Client Code"], errors='coerce').fillna(0).astype(int)
        df["Xero Entity Name"] = df["Xero Entity Name"].fillna("TBC")
        
        # Lower caps
        df['Xero Entity Name'] = df['Xero Entity Name'].str.lower()
    
        # create auxilary lists
        churn_df = df.loc[(df["Xero Entity Name"] == self.CHURNED.lower() )].drop(columns=["Xero Entity Name"])
        tbc_df = df.loc[(df["Xero Entity Name"] == "TBC".lower() )].drop(columns=["Xero Entity Name"])
        self.churn_dict: Dict[str,dict] = churn_df.set_index("Client Name").T.to_dict()
        self.tbc_dict: Dict[str,dict] = tbc_df.set_index("Client Name").T.to_dict()

        # create clean dict
        clean_df = df.loc[(df["Xero Entity Name"] != "TBC" ) & (df["Xero Entity Name"] != self.CHURNED)]
        self.entity_key_dict: Dict[str,List] = clean_df.set_index("Xero Entity Name").T.to_dict('list')
        self.code_key_dict: Dict[str,List]= clean_df.set_index("Client Code").T.to_dict('list')
    
    def find_client_name_by_entity_name(self, entity_name:str) -> List|None:
        """Returns client name by searching using entity name

        Args:
            entity_name (str): _description_

        Returns:
            str: _description_
        """
        record:List = self.entity_key_dict.get(entity_name,None)
        if record != None:
            return record
        # return client_name, client_code
        return [None, None]
            
    def find_client_name_by_code(self, customer_code:str) -> str|None:
        """Returns client name by searching using customer code

        Args:
            customer_code (str): _description_

        Returns:
            str: _description_
        """
        try:
            record: str = self.code_key_dict.get(int(customer_code),None)
            if record != None:
                return record[0]
        except Exception as e:
            pass
        return None

    def force_find_client_name(self, description_field:str) -> List:
        """Brute force search client name in description field.
        Try each index in clean list, then CHURNED

        Returns:
            str: _description_
        """
        # Check if the description field is valid
        if pd.isnull(description_field):
            return None, None
        
        # Test each entity name in the dict
        for key in self.entity_key_dict:
            if key in description_field:
                # key subset of description
                # return client name from list
                return self.entity_key_dict[key]
        # Test all churned item
        for key in self.churn_dict:
            if key in description_field:
                # key subset of description
                return key
        return None, None

    def populate_client_name(self) -> pd.DataFrame:
        # for each row in dataframe
        for index, row in self.pl_df.iterrows():
            if row["Client name"] != "":
                # if already filled, dont waste time
                continue

            found:bool = False
            # if there exist a contact
            if not pd.isnull(row["Contact"]):
                client_name, customer_code = self.find_client_name_by_entity_name(row["Contact"])
                if client_name:
                    # self.pl_df.iloc[index]["Client name"] = client_name
                    # row["Client name"] = client_name
                    self.pl_df.loc[self.pl_df["Contact"]==row["Contact"], "Client name"] = client_name
                    self.pl_df.loc[self.pl_df["Contact"]==row["Contact"], "Client code"] = customer_code
                    continue

            # if there exist a reference number
            if not pd.isnull(row["Reference"]):
                # get the first number
                customer_code:str = row["Reference"].split("-")[0]
                customer_code = Profit_and_Loss.clean_customer_code(customer_code)
                # check number against the client list
                client_name = self.find_client_name_by_code(customer_code)
                if client_name:
                    self.pl_df.loc[self.pl_df["Reference"]==row["Reference"], "Client name"] = client_name
                    self.pl_df.loc[self.pl_df["Reference"]==row["Reference"], "Client code"] = customer_code
                    continue
                
            # if all fails
            # then for each valid xero entity name
            # check if it exists within the description
            # if true, then update the client field as such
            client_name, customer_code = self.force_find_client_name(row["Description"])
            if client_name != None:
                self.pl_df.loc[self.pl_df["Description"]==row["Description"], "Client name"] = client_name
                self.pl_df.loc[self.pl_df["Description"]==row["Description"], "Client code"] = customer_code
                self.pl_df.loc[self.pl_df["Description"]==row["Description"], "To Check"] = "True"
        

    def save_to_csv(self, filename) -> None:
        np.savetxt("compiled.csv", 
           self.pl_df,
           delimiter =",", 
           fmt ='%s')
    def save_to_excel(self, filename) -> None:
        self.pl_df.to_excel(filename)
    
    def save_dict_mapping(self) -> None:
        import csv
        with open("entitymap.csv", 'w') as f:  # You will need 'wb' mode in Python 2.x
            w = csv.DictWriter(f, self.entity_key_dict.keys())
            w.writeheader()
            w.writerow(self.entity_key_dict)
        
    def start(self) -> None:
        # update dataframe
        self.populate_client_name()
        self.restore_orignal_columns()
        self.save_to_excel("compiled.xlsx")
        # Note that description and contact is not preserved

t = Profit_and_Loss()
t.start()