import mysql # type: ignore
import mysql.connector # type: ignore
import mysql.connector.abstracts # type: ignore
import mysql.connector.cursor # type: ignore
import pandas as pd # type: ignore
import re
import numpy as np
from logger import Logger
from type import *
from typing import Iterable

ERROR = "MySQL Connection Failed! Please try again"
NO_ID = "HR Emp Code column was not found!"

class CreateTable:
    SUCCESS = "Table Generated Successfully"
    EXISTS = "Table Already Exists"
    COLUMNS_MISMATCH = "Table exists, but columns do not match to those in database"
    ERROR = ERROR
    NO_ID = NO_ID


class UpdateTable:
    ERROR = ERROR
    COLUMNS_MISMATCH = "Table exists, but columns do not match to those in database"
    SUCCESS = "Records were successfully inserted or updated!"
    NO_ID = NO_ID

class DeleteTable:
    ERROR = ERROR
    SUCCESS = "Table was successfully deleted!"
    TABLE_NOT_FOUND = "Table does not exists!"

TABLE_FORMAT = r'^(somaiya|svv)_(teaching|nonteaching|temporary|svv)_(jan|feb|mar|apr|may|jun|jul|aug|sept|oct|nov|dec)_(\d{4})\Z' # insti_type_month_year

def sanitize_column(txt:str) -> str:
    """ Sanitizing identifiers """
    def func(txt:str): return str(txt).replace('`','``').replace('\n','').strip()
    return f"`{func(txt)}`"

def sanitize_value(txt:str) -> str:
    """ Sanitizing values """
    return f"""'{str(txt).replace("'","''").replace('\n','').strip()}'"""

def dataRefine(data:pd.DataFrame) -> None:
    """ refines columns for sql in place """
    data.rename(columns={col: str(col).strip().replace('\n','') for col in data.columns}, inplace=True)

def cleanData(val:str|int|float) -> str:
    
    try:
        val = int(val) if ((type(val)==np.int64) or (type(val)==np.float64) or (type(val)==float) or (type(val)==int)) else str(val)
    except:
        val = str(val)

    if type(val) is str:
        txt = re.findall(r"(\d{4})[-,/](\d{1,2})[-,/](\d{1,2})", str(val))
        
        if txt:
            val='-'.join([i for i in txt[0][::-1]])

    return str(val).replace('"',"'")


def mapping(pd_columns: Iterable[str], columns:str) -> str | None:
    column = columns.lower()
    
    for i in sorted(pd_columns, key=len): 
        if(column in i.lower()): return i
        
    return None

def check_column(col:str, pd_Data:pd.DataFrame) -> NullStr:
    pred_col = mapping(pd_Data.columns, col)

    if (pred_col is not None and (data := pd_Data[pred_col].to_numpy(str))):
        return None if (not data) else data[0]
    else: 
        return None
    
class Database():
    def __init__(self, error_logger:Logger) -> None:
        self.db: Optional[mysql.connector.abstracts.MySQLConnectionAbstract | mysql.connector.pooling.PooledMySQLConnection] = None
        self.logger = error_logger
        
        
    @staticmethod
    def getTableName(month: MonthList, year:int, insti: InstituteList, type: TypeList) -> str:
        return sanitize_column(f'{insti.lower()}_{type.lower()}_{month.lower()}_{year}')

    def connectDatabase(self, host:str, user:str, password:str, database:str):
        try:
            self.db = mysql.connector.connect( 
                        host=host,
                        user=user,
                        password=password,
                        database=database
                    )
                
            self.add_mysql_info(f'Connected to Database {database}')
        except Exception as e:
            self.add_mysql_error(self.logger.get_error_info(e))
            
        return self
    
    def isConnected(self):
        
        if(self.db is None):
            return False
        
        try:
            return bool(self.db.is_connected())
        except Exception as e:
            self.add_mysql_error(self.logger.get_error_info(e))
        
        return False
    
    def column_check(self, month: MonthList, year:int, columns: Iterable[str], insti: InstituteList, type: TypeList) -> bool:
        
        db_columns = set(self.getColumns(month, year, insti, type))
        seen: set[str] = set()
        
        for col in columns:
            if((col in seen) or (col not in db_columns)): return False
            seen.add(col)

        return True
    
    def createData(self, month: MonthList, year:int, columns:list[str], insti: InstituteList, type: TypeList) -> str:
        """ create a table from month and year if it does not exist """
        
        columns = sorted(columns)
        
        code_col = mapping(columns="hr emp code", pd_columns=columns)

        if(not code_col): 
            return CreateTable.NO_ID
        
        if(self.db is None): return CreateTable.ERROR
        
        try:
            
            with self.db.cursor() as cursor:
                table_name = self.getTableName(month, year, insti, type)
                sql = f"""CREATE TABLE {table_name}(
                    {','.join(
                        [ f'{sanitize_column(col)} VARCHAR(225) PRIMARY KEY' if (col==code_col) else f'{sanitize_column(col)} VARCHAR(225)'
                            for col in columns
                        ]
                        )
                    }
                )"""
                
                cursor.execute(sql)
                self.add_mysql_info(f'Created table {table_name}')
                self.db.commit()
                return CreateTable.SUCCESS

        except mysql.connector.errors.ProgrammingError as e:
            self.add_mysql_error(self.logger.get_error_info(e))
            
            if(sorted(self.getColumns(month,year,insti,type))!=sorted(columns)):
                return CreateTable.COLUMNS_MISMATCH
            else:
                return CreateTable.EXISTS
    
        except Exception as e:
            self.add_mysql_error(self.logger.get_error_info(e))
            return CreateTable.ERROR

    def updateData(self, data:pd.DataFrame, month: MonthList, year:int, insti: InstituteList, type: TypeList) -> str:
        """ updates existing data or inserts new data """
        
        id = mapping(pd_columns=data.columns, columns='HR EMP CODE')
        
        columns = {j:i for i,j in enumerate(data.columns)}
        
        if (not self.column_check(month, year, data.columns, insti, type)):
            return UpdateTable.COLUMNS_MISMATCH
        
        if (not id): 
            return UpdateTable.NO_ID
        
        if(self.db is None): return UpdateTable.ERROR
        
        keys = ','.join(map(sanitize_column,list(columns.keys())))
        table_name = self.getTableName(month, year, insti, type)
        
        try:
            cursor = self.db.cursor()
            
            with self.db.cursor() as cursor:
                for row in data.itertuples(index=False):
                    row_data = {col:cleanData(row[columns[col]]) for col in data.columns}
                    query =','.join([f"{sanitize_column(col)}={sanitize_value(data[col])}" for col in row_data]) # type: ignore
                    values = ','.join(map(sanitize_value,list(row_data.values())))

                    try:
                        cursor.execute(f"INSERT INTO {table_name} ({keys}) VALUE ({values});")    
                        self.add_mysql_info(f"Inserting data into {table_name}")

                    except mysql.connector.errors.IntegrityError as e:
                        self.add_mysql_error(self.logger.get_error_info(e))
                        
                        try:
                            cursor.execute(f"UPDATE {table_name} SET {query} WHERE {sanitize_column(id)}={sanitize_value(data[id])};") # type: ignore
                            self.add_mysql_info(f"Updating data into {table_name}")
                            
                        except Exception as f:
                            self.add_mysql_error(self.logger.get_error_info(f))
                            return UpdateTable.ERROR
                    
                    self.db.commit()
            
        except Exception as g:
            self.add_mysql_error(self.logger.get_error_info(g))
            return  UpdateTable.ERROR
        
        return UpdateTable.SUCCESS
    
    def dropTable(self, month: MonthList, year:int, insti: InstituteList, type: TypeList) -> str:
        """ Drops the table """
        
        if(self.db is None): return DeleteTable.ERROR
        
        table_name = self.getTableName(month, year, insti, type)
        
        try:
            with self.db.cursor() as cursor:
                cursor.execute(f'drop table {table_name}')
                self.add_mysql_info(f'Deleting data from {table_name} (Hope you have backup!)')
                self.db.commit()
                return DeleteTable.SUCCESS
        
        except mysql.connector.ProgrammingError as f:
            self.add_mysql_error(self.logger.get_error_info(f))
            return DeleteTable.TABLE_NOT_FOUND
        
        except Exception as e:
            self.add_mysql_error(self.logger.get_error_info(e))
            return DeleteTable.ERROR
        
    def showTables(self) -> dict[str,dict[str,dict[str,set[str]]]]:
        """ Show all tables """
        
        memo: dict[str, dict[str, dict[str, set[str]]]] = {}
        tables: list[str] = []
        
        if(self.db is None): return memo
        
        try:
            with self.db.cursor() as cursor:
                cursor.execute('SHOW TABLES')
                tables = [str(col[0]) for col in cursor.fetchall()] # type: ignore
                self.add_mysql_info(f'Fetching table(s) info')
            
        except Exception as e:
            self.add_mysql_error(self.logger.get_error_info(e))    
            return memo

        for table in tables:
            
            if (table_data := re.match(TABLE_FORMAT,table)) is not None:
                insti, type, month, year = table_data.groups()
                
                if insti not in memo: 
                    memo[insti] = {}
                if type not in memo[insti]: 
                    memo[insti][type] = {}
                if year not in memo[insti][type]: 
                    memo[insti][type][year] = set()
        
                memo[insti][type][year].add(month)

            else:
                expected_format = r"Expected table name format as '^(somaiya|svv)_(teaching|nonteaching|temporary|svv)_(jan|feb|mar|apr|may|jun|jul|aug|sept|oct|nov|dec)_(\d{4})\Z"
                self.add_mysql_error(f"Unexpected table name format: {table}. {expected_format}")

        return memo

    def getColumns(self, month: MonthList, year:int, insti: InstituteList, type: TypeList) -> list[str]:
        """ fetches all columns of a table """
        if(self.db is None): return []

        try:
            
            with self.db.cursor() as cursor:
            
                cursor.execute(f'desc {sanitize_column(f"{insti.lower()}_{type.lower()}_{month.lower()}_{year}")}')
                self.add_mysql_info(f'Checking table {sanitize_column(f"{insti.lower()}_{type.lower()}_{month.lower()}_{year}")} info')
                return [str(col_data[0]) for col_data in cursor.fetchall()] # type: ignore
            
        except Exception as e:
            self.add_mysql_error(self.logger.get_error_info(e))    
            return []

    def fetchAll(self, month: MonthList, year:int, insti: InstituteList, type: TypeList) -> pd.DataFrame|None:
        """ fetches all data from table month_year """
        
        data = pd.DataFrame()
        
        if(self.db is None): return data
        
        table_name = self.getTableName(month, year, insti, type)
        
        try:
            columns = self.getColumns(month,year,insti,type)
            with self.db.cursor() as cursor:
                cursor.execute(f"SELECT * FROM {table_name}")
                self.add_mysql_info(f"Fetching data from table {table_name}")
                
                result = cursor.fetchall()
            
                return pd.DataFrame(result, columns=columns,dtype=str)
        
        except Exception as e:
            self.add_mysql_error(self.logger.get_error_info(e))    
            return data
        
    def endDatabase(self):
        """ end database. RIP """
        if self.db is None: return self
        
        try:
            self.db.close()
            self.add_mysql_info(f'Closing connection o7')
            
        except Exception as e:
            self.add_mysql_error(self.logger.get_error_info(e))    
            
        return self
    
    def add_mysql_error(self, msg:str):
        self.logger.write_error(msg,'MySQL')
            
    def add_mysql_info(self, msg:str):
        self.logger.write_info(msg,'MySQL')        