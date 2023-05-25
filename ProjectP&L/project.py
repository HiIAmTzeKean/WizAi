import pandas as pd
import csv
import re
import os
from pathlib import Path
from typing import List, Dict, Any
from dataclasses import dataclass, fields

@dataclass
class Expenses:
    currency:str=""
    airticket_international:float=0
    airticket_national:float=0
    accomodation_international:float=0
    accomodation_domestic:float=0
    allowance:float=0
    transport:float=0
    meal:float=0
    def to_list(self) -> List[Any]:
        return [
            self.airticket_international,
            self.airticket_national,
            self.accomodation_international,
            self.accomodation_domestic,
            self.allowance,
            self.transport
        ]
        
@dataclass
class DataRow:
    key:str=""
    unique_id:str=""
    salesman:str=""
    department:str=""
    currency:str=""
    month:str=""
    airticket_international:float=0
    airticket_national:float=0
    accomodation_international:float=0
    accomodation_domestic:float=0
    allowance:float=0
    transport:float=0
    meal:float=0
    def to_list(self) -> List[Any]:
        return [
            self.key,
            self.unique_id,
            self.month,
            self.airticket_international,
            self.airticket_national,
            self.accomodation_international,
            self.accomodation_domestic,
            self.allowance,
            self.transport,
            self.meal,
            self.currency,
            self.salesman,
            self.department
        ]
    def get_header():
        return [
            "key",
            "unique_id",
            "month",
            "airticket_international",
            "airticket_national",
            "accomodation_international",
            "accomodation_domestic",
            "allowance",
            "transport",
            "meal",
            "currency",
            "salesman",
            "department"
        ]

class ProductLine():
    def __init__(self,unique_id:str,client_name:str,allocation:float,currency:str) -> None:
        self.unique_id:str = unique_id
        self.client_name:str = client_name
        self.allocation:float = allocation
        self.currency:str = currency
        self.expense = ProductLine.create_expense_type()
        
    def create_expense_type() -> Expenses:
        return Expenses()
    
    def create_expense(self,costs:Expenses) -> None:
        """Creates expense object for storage of expenses

        Args:
            costs (Expenses): _description_
        """
        for field in fields(Expenses):
            if field.name == "currency" or field.name == "meal":
                continue
            self.expense.__setattr__(field.name, getattr(costs, field.name) * self.allocation)
            
    def update_expense(self, costs:Expenses, allocation:float) -> None:
        for field in fields(Expenses):
            if field.name == "currency" or field.name == "meal":
                continue
            self.expense.__setattr__(field.name, getattr(costs, field.name) * self.allocation
                                                + self.expense.__getattribute__(field.name))
            
    def update_meal(self, amount) -> None:
        self.expense.meal += amount
    
    def __add__(self, o:Any):
        if o.__class__!=ProductLine:
            return self
        o: ProductLine = o
        for field in fields(Expenses):
            if field.name == "currency":
                continue
            self.expense.__setattr__(field.name,
                    self.expense.__getattribute__(field.name)+o.expense.__getattribute__(field.name))
        return self
    
    def format_for_csv(self, salesman:str, department:str, month:str) -> DataRow:
        row = DataRow()
        for field in fields(Expenses):
            row.__setattr__(field.name,self.expense.__getattribute__(field.name))
        row.salesman = salesman
        row.unique_id = self.unique_id
        row.department = department
        row.month = month
        row.currency = self.currency
        salesman_clean = salesman.replace(",","").replace(" ","").strip()
        row.key = f"2023_{month}_{self.unique_id}_{salesman_clean}"
        return row

class TravelExpense():
    def __init__(self,currency:str,
                 airticket_international:str,
                 airticket_national:str,
                 accomodation_international:str,
                 accomodation_domestic:str,
                 allowance:str,
                 transport:str) -> None:
        self.currency:str=currency
        self.airticket_international:float=TravelExpense.convert_to_float(airticket_international)
        self.airticket_national:float=TravelExpense.convert_to_float(airticket_national)
        self.accomodation_international:float=TravelExpense.convert_to_float(accomodation_international)
        self.accomodation_domestic:float=TravelExpense.convert_to_float(accomodation_domestic)
        self.allowance:float=TravelExpense.convert_to_float(allowance)
        self.transport:float=TravelExpense.convert_to_float(transport)
        
    def convert_to_float(item) -> float:
        try:
            return float(item)
        except:
            print(f"unable to convert {item} to float")
            return 0
        
    def get_expense(self) -> Expenses:
        return Expenses(self.currency,
                self.airticket_international,
                self.airticket_national,
                self.accomodation_international,
                self.accomodation_domestic,
                self.allowance,
                self.transport)
        
class Form():
    unique_ID_REGEX = r"^([0-9]*-[a-zA-Z]*[0-9]*-[0-9]*|[0-9]*)$"
    def __init__(self,target_wb:str,target_month:str) -> None:
        self.form_name:str
        self.target_wb = target_wb
        self.target_month = target_month
        self.salesperson:str = ""
        self.department:str = ""
        self.country:str =""
        
    def valid_ID(string:str):
        if re.match(Form.unique_ID_REGEX, string):
            return True
        return False

    def process_form(self,df:pd.DataFrame):
        try:
            df = self.get_meta_data(df)
            df,expense = self.get_travel_expense(df)
        except KeyError as e:
            # No more rows, so end program
            return
        
        while df.iloc[0,0]!="Unique ID":
            df = df.iloc[1:, :]
        
        # Skip the header
        df = df.iloc[1:, :]
        # Update header
        df.columns = ["unique id","client name", "project name", "project allocation", "project allocation (source)", "remark", "approve"]
        allocation:Dict[str:ProductLine] = dict()
        
        row = df.iloc[0]
        # Note that unique ID can be replaced with customer code (XXX) type
        while pd.notnull(row["unique id"]) and Form.valid_ID(str(row["unique id"])):
            uniqueID = str(str(row["unique id"]))
            # Check if the product line exists
            try:
                item:ProductLine = allocation.get(uniqueID,None)
                if item==None:
                    # Product line does not exist and has to be created
                    item:ProductLine = ProductLine(uniqueID,
                                                row["client name"],
                                                row["project allocation"],
                                                expense.currency)
                    item.create_expense(expense)
                    allocation[uniqueID] = item
                else:
                    item.update_expense(expense,row["project allocation"])
            except Exception as e:
                print(f"{uniqueID} already exists before!")
            df = df.iloc[1:, :]
            row = df.iloc[0]
        
        while df.iloc[0,0]!="Unique ID":
            df = df.iloc[1:, :]
            
        # Skip the header
        df = df.iloc[1:, :]
        # Update header
        df.columns = ["unique id","client name", "project name", "percentage", "amount", "date", "approve"]
        ## to modify, no formula
        row = df.iloc[0]
        while pd.notnull(row["unique id"]) and Form.valid_ID(str(row["unique id"])):
            uniqueID = str(row["unique id"])
            try:
                item:ProductLine = allocation.get(uniqueID,None)
                if item==None:
                    item:ProductLine = ProductLine(uniqueID,
                                        row["client name"],
                                        0,
                                        expense.currency)
                    allocation[uniqueID]=item
                item.update_meal(row["amount"])
            except Exception as e:
                print(e)
            df = df.iloc[1:, :]
            row = df.iloc[0]
        
        return df,allocation
    
    def get_meta_data(self,df) -> pd.DataFrame:
        try:
            while df.iloc[0,0]!="Requester's name":
                df = df.iloc[1:, :]
            
            # Skip the header
            df = df.iloc[1:, :]
            if not self.salesperson:
                self.salesperson: str = df.iloc[0,0]
                self.department: str = df.iloc[0,1]
                self.country: str = df.iloc[0,3]
            return df
        except Exception as e:
            raise EOFError
    
    def get_travel_expense(self,df):
        while df.iloc[0,0]!="Type":
            df = df.iloc[1:, :]
        
        # Skip the header
        df = df.iloc[1:, :]
        expense = TravelExpense(currency=df.iloc[0,1],
                        airticket_international=df.iloc[0,2],
                        airticket_national=df.iloc[1,2],
                        accomodation_international=df.iloc[2,2],
                        accomodation_domestic=df.iloc[3,2],
                        allowance=df.iloc[4,2],
                        transport=df.iloc[5,2])
        return (df,expense)
        
    def process_forms(self) -> List[DataRow]:
        df = pd.read_excel(self.target_wb,sheet_name=self.target_month)
        expense:List[Dict] = []
        while True:
            try:
                df, allocation = self.process_form(df)
                expense.append(allocation)
            except Exception as e:
                print(e)
                break
        expense = Form.collate_expenses(expense)
        
        return self.format_for_csv(expense)
    
    def format_for_csv(self,list_item:List[ProductLine]) -> List[DataRow]:
        for i in range(len(list_item)):
            list_item[i] = list_item[i].format_for_csv(self.salesperson,self.department,self.target_month)
        return list_item
    
    def collate_expenses(expense:List[Dict]) -> List[ProductLine]:
        result:Dict = {}
        for allocation in expense:
            result.update({k: allocation.get(k, None) + result.get(k, None) for k in set(allocation)})
        return list(result.values())

class ExcelCreator():
    def __init__(self) -> None:
        pass
    def start():
        target_month = "May"
        output_filename = "output.csv"
        collated_records:List[DataRow] = []

        dir_path = os.path.dirname(os.path.realpath(__file__))
        files = Path(dir_path).glob('*.xlsx')
        for file in files:
            collated_records.extend(ExcelCreator.process_file(file,target_month))
            
        output_file_dir = os.path.join(dir_path,output_filename)
        
        ExcelCreator.write_to_csv(output_file_dir,collated_records)
    
    def process_file(target_wb,target_month):
        f = Form(target_wb,target_month)
        return f.process_forms()

    def write_to_csv(filename:str, list_item:List[DataRow]):
        with open(filename,'w',newline='') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(DataRow.get_header())
            for item in list_item:
                csvwriter.writerow(item.to_list())
            
ExcelCreator.start()