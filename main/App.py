import base64
import json
from pathlib import Path
from type import *
import io, os, sys, re, pdfkit, pyperclip, msoffcrypto, gc, traceback # type: ignore
import customtkinter as ctk # type: ignore
import tkinter as tk # type: ignore
import pandas as pd # type: ignore
from mail import Mailing, AsyncMailing, AsyncMessage, MIMEMultipart
from logger import Logger
from threading import Thread, excepthook
import tkinter.messagebox as tkmb
from default import SVG_ICON, TEMPLATE
from multiprocessing import Process, Queue, freeze_support
from tkinter import filedialog, scrolledtext, Scrollbar, messagebox 
from database import dataRefine, Database, mapping, CreateTable, UpdateTable, DeleteTable
from PIL import Image, ImageTk
import pdfkit
from parser import PDFTemplate
from asyncio import gather, run, to_thread
from copy import deepcopy
from creds import PROD_CREDS, TEST_CREDS, DB_CREDS

IS_EXE = True
""" when building .exe set to true """

IS_DEBUG = False 
""" for testing set to true """

APP_PATH = __file__ if not IS_EXE else sys.executable

MYSQL_CRED =  DB_CREDS 
""" your mysql creds for testing (works if is_debug == True) """

MAIL_CRED = PROD_CREDS if not IS_DEBUG else TEST_CREDS

FILE_REGEX = r'employee_(\w+|\d+).pdf\Z'

CODE_COL = "__id__"

TEMPLATE_SHEET = ["Personal Left", "Personal Right", "Earning", "Deductions", "Salary Left", "Salary Right"]
TEMPLATE_COLUMN = ["Name", "Column"]

TYPE = type

def email_check(x: str):
    """ a helper function for email validation """
    if re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\Z', x):
        return True
    else:
        return False
        
def year_check(x: str): 
    """ a helper function for year validation """
    if re.match(r'^(\d){4}\Z',x):
        return True 
    else:
        return False
    
def text_clean(x: str):
    return str(x).replace('\n','').strip()

def file_clean(x: str):
    
    def _clean(string: str, not_allowed: str) -> str:
        deny = set(not_allowed)
        new_string: str = ""
        
        for i in string:
            if(i not in deny): new_string += i
        return "none" if (not new_string) else new_string
    
    return _clean(x, """<>:"/\\?*'\n""")

def checkColumns(present_col: list[str], needed_col: list[str]) -> bool:
    seen = set()
    for i in present_col:
        if (mapping(needed_col, i) is not None):
            seen.add(i)
    return len(needed_col) == len(seen)

ERROR_LOG = Logger(Path(APP_PATH).parent)
PDF_TEMPLATE = PDFTemplate(Path(APP_PATH).parent,ERROR_LOG)

MIN_TEXT_SIZE:int = 12
MAX_TEXT_SIZE:int = 25

COLOR_SCHEME = {"fg_color": "white","button_color": "dark red","text_color": "black","combo_box_color": "white"}

ctk.set_appearance_mode("light")

class BaseTemplate():
    """ Base Template for all tkinter frames """
    
    process: Process | None = None
    """ All frames will have 1 common process to run big task """
    
    thread: Thread | None = None
    """ All frames gets 1 common thread to process gui changes """
    
    QUEUE: Queue = Queue()
    """ A queue for communication between thread and process """
    
    stop_flag: bool = False
    """ A stop flag for thread to stop """
    
    data: pd.DataFrame | dict[str, pd.DataFrame] | None = None
    """ A shared data storage """
    
    def __init__(self, outer: 'App') -> None:
        self.outer = outer
        self.visible: bool = False
        self.to_disable: list[ctk.CTkButton|ctk.CTkOptionMenu|ctk.CTkEntry] = []
        self.frame: ctk.CTkScrollableFrame = ctk.CTkScrollableFrame(master=self.outer.APP, fg_color=COLOR_SCHEME["fg_color"])
        
    def appear(self) -> None:
        """ Add frame to the tkinter app """
        
        if(not self.visible):
            self.outer.start_frame.pack(pady=0)
            self.frame.pack(pady=5, padx=20, fill='both', expand=True)
            self.outer.credit_frame.pack(pady=5, padx=20, fill='both')
            self.outer.end_frame.pack(pady=0)
            self.visible = True

    def hide(self) -> None:
        """ Hides the current frame """
        
        if(self.visible):        
            self.frame.pack_forget()
            self.outer.start_frame.pack_forget()
            self.outer.end_frame.pack_forget()
            self.outer.credit_frame.pack_forget()
            self.visible = False
            
    def cancel_thread(self) -> None:
        """ Basic function to cancel/quit out of process and thread (using flags) """
        
        if(self.thread and self.thread.is_alive()):
            self.stop_flag = True
            
        if(self.process and self.process.is_alive()):
            self.process.terminate()
            
        self.clear_queue()

            
    def can_start_thread(self) -> bool:
        """ Check if threading can begin """
        
        can_we_start = bool((self.process is None) or (self.process and (not self.process.is_alive())) and ((self.thread is None) or (self.thread and (not self.thread.is_alive()))))
        
        if(can_we_start): self.stop_flag = False
        
        return can_we_start
    
    def get_widgets_to_disable(self):
        """ Finds widgets to disable """
        
        all_attribute = dir(self)
        
        for attr in (all_attribute):
            
            if("quit" in attr): continue
            
            widget = self.__getattribute__(attr)
            
            if((isinstance(widget, ctk.CTkButton) or isinstance(widget, ctk.CTkOptionMenu) or isinstance(widget, ctk.CTkEntry))):
                yield widget

    def clear_queue(self):
        """ Clear the queue """
        while not self.QUEUE.empty(): self.QUEUE.get()
    
    def clear_data(self, hard = True):
        """ Clear Shared Data """
        
        if(hard): BaseTemplate.data = None
        erased = gc.collect()
        ERROR_LOG.write_info(f"{erased} Variables Cleared")
    
    def switch_screen(self, app: type["BaseTemplate"]):
        self.hide()
        self.outer.CHILD[app.__name__].appear() # type: ignore
    
class App:
    """ Main class representing the application """
    
    APP:ctk.CTk = ctk.CTk()    
    TOGGLE:dict[str, tuple[str, ...]] = {'Somaiya':('Teaching','Non-Teaching','Temporary'),'SVV':('svv',)}
    
    CRED = DB_CRED(host="", user="", password="", database="")
    MONTH:dict[str,int] = {"jan":1, "feb":2, "mar":3, "apr":4, "may":5, "jun":6, "jul":7, "aug":8, "sept":9, "oct":10, "nov":11, "dec":12}
    DB = Database(ERROR_LOG)
    
    credit_frame = ctk.CTkFrame(master=APP,height=10,fg_color=COLOR_SCHEME['fg_color'])
    start_frame = ctk.CTkLabel(text='',master=APP,height=10)
    end_frame = ctk.CTkLabel(text='',master=APP,height=10)

    def __init__(self, width:int, height:int, title:str) -> None:
        self.APP.geometry(f"{width}x{height}")
        self.APP.title(title)
        
        self.CHILD:dict[str, 
                        Union[
                            UploadData,
                            FileInput,
                            DataView,
                            MySQLLogin,
                            SendMail,
                            SendBulkMail,
                            Login,
                            Interface,
                            MailCover,
                            DataPreview,
                            DeleteView,
                            DataPeek,
                            TemplateInput,
                            TemplateGeneration
                        ]] = {
            UploadData.__name__: UploadData(self),
            FileInput.__name__: FileInput(self), 
            DataView.__name__: DataView(self),
            MySQLLogin.__name__: MySQLLogin(self),
            SendMail.__name__: SendMail(self),
            SendBulkMail.__name__: SendBulkMail(self),
            Login.__name__: Login(self),
            Interface.__name__: Interface(self),
            MailCover.__name__: MailCover(self),
            DataPreview.__name__: DataPreview(self),
            DeleteView.__name__: DeleteView(self),
            DataPeek.__name__: DataPeek(self),
            TemplateInput.__name__: TemplateInput(self),
            TemplateGeneration.__name__: TemplateGeneration(self),
        }
        
    def start_app(self):
        """ To start application run this """
        
        self.CHILD[Login.__name__].switch_screen(Login)
        self.credits()
        self.APP.mainloop()
        
    def exit_app(self):
        try:
            self.APP.destroy()
        except Exception as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
        finally:
            ERROR_LOG.write_info('User Logged Out')
    
    def credits(self) -> None:
        img_data = Image.open(io.BytesIO(base64.b64decode(SVG_ICON)))
        img = ImageTk.PhotoImage(img_data)
        img = ctk.CTkImage(img_data,img_data,(243,65))
        ctk.CTkLabel(master=self.credit_frame, image=img, text='',text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=5,padx=10,side='left')
        ctk.CTkLabel(master=self.credit_frame, text='',text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=5,padx=10,side='right')
        ctk.CTkLabel(master=self.credit_frame , text="Under Guidance of: Dr. Sarita Ambadekar and Dr. Abhijit Patil", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=5,padx=10)
        ctk.CTkLabel(master=self.credit_frame , text="First developed by: Raj More, Pranav Lohar, Aryan Mandke.", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=5,padx=10)
        ctk.CTkLabel(master=self.credit_frame , text="Department of Computer Engineering", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=5,padx=10)
    

class PDFGenerator:
    """ Class to handle everything related to PDF Generation """
    
    @staticmethod
    def _generate_pdf(pdf_file_path:Path, html_content:str) -> None:
        """ Generates the pdf from html """

        pdfkit.from_string(
            html_content, 
            str(pdf_file_path),
            options={
                'page-size': 'A4',
                'margin-top': '0.5in',
                'margin-right': '0.5in',
                'margin-bottom': '0.5in',
                'margin-left': '0.5in',
                'no-outline': None
            }
        )

    @staticmethod
    def generate_one_pdf(name: str, html_content:str, file_path:Path) -> tuple[bool,str]:
        """ Wrapper Function to generate one pdf  """
        
        try:
            where = Path(file_path).parent

            filename = f"employee_{name}.pdf"

            if (re.match(FILE_REGEX,filename)):
                pdf_file = where.joinpath(filename)
                PDFGenerator._generate_pdf(pdf_file, html_content)
                return (True,"PDF Generation Successful")
            else:
                ERROR_LOG.write_info(f"{filename} does not match pattern '{FILE_REGEX}'")
                return (False,f"{filename} does not match pattern '{FILE_REGEX}'")
                
        except Exception as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
            
            return (False,f"An error occurred while generating the PDF: {str(e)}")
            

    @staticmethod
    async def generate_many_pdf(name: str, html_content:str, where:Path) -> tuple[bool,str]:
        """ Wrapper Function to generate multiple pdfs """
        
        try:
            filename = f"employee_{file_clean(name)}.pdf"
            if (re.match(FILE_REGEX,filename)):
                pdf_file = where.joinpath(filename)
                await to_thread(PDFGenerator._generate_pdf, pdf_file, html_content)
                return (True,"PDF Generation Successful")
            else:
                ERROR_LOG.write_info(f"{filename} does not match pattern '{FILE_REGEX}'")
                return (False,f"{filename} does not match pattern '{FILE_REGEX}'")
                
        except Exception as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
            return (False,f"An error occurred while generating the PDF: {str(e)}")

class GUI_Handler:
    """ Handles GUI """
    
    @staticmethod
    def view_excel(data: pd.DataFrame, text_excel:scrolledtext.ScrolledText) -> None:
        """ Changes the text on the text excel thing  """
        
        text = [[str(i) for i in data.columns]]
        col_widths = {i:len(str(j)) for i,j in enumerate(data.columns)}
        
        for row in data.itertuples(index=False):
            row_data = []
            
            for idx, cell in enumerate(row):
                row_data.append(cell)
                col_widths[idx] = max(col_widths[idx], len(str(cell)))
                
            text.append(row_data)
        
        formatted_data = "\n".join([" " + " | ".join([f"{text_clean(cell):^{col_widths[i]}}" for i, cell in enumerate(row)]) + " " for row in text])
        
        text_excel.delete(1.0, tk.END)
        text_excel.insert(tk.END, formatted_data)
        
        text_excel.tag_config("highlight", font=("Courier", 12, "bold"), foreground="yellow")
        text_excel.tag_add("highlight", "1.0", "1.end")
    
    @staticmethod
    def clear_excel(text_excel: scrolledtext.ScrolledText) -> None:
        text_excel.delete(1.0, tk.END)
    
    @staticmethod    
    def change_text_font(text_excel: scrolledtext.ScrolledText, size:int = 15) -> None:
        """ Change the size of text """
        text_excel.configure(font=("Courier", size))
    
    @staticmethod
    def lock_gui_button(buttons: Iterator[ctk.CTkButton] | list[ctk.CTkButton]) -> None:
        """ Disable buttons """
        for button in buttons:
            button.configure(state="disabled")

    @staticmethod
    def unlock_gui_button(buttons: Iterable[ctk.CTkButton]) -> None:
        """ Enable buttons """
        for button in buttons:
            button.configure(state="normal")

    @staticmethod
    def change_file_holder(file_path_holder:ctk.CTkEntry, new_path:str) -> None:
        """ Change the file path holders """
        file_path_holder.delete(0, tk.END)
        file_path_holder.insert(0,new_path)
        file_path_holder.configure(width=7*len(new_path))
        
    @staticmethod
    def place_after(anchor:ctk.CTkBaseClass, to_place:ctk.CTkBaseClass, padx:int = 10, pady:int = 10) -> None:
        """ Places widget after the anchor widgets """
        to_place.pack_configure(after=anchor,padx=padx,pady=pady)
    
    @staticmethod
    def place_before(anchor:ctk.CTkBaseClass, to_place:ctk.CTkBaseClass, padx:int = 10, pady:int = 10) -> None:
        """ Places widget before the anchor widgets """
        to_place.pack_configure(before=anchor,padx=padx,pady=pady)
        
    @staticmethod
    def remove_widget(this_widget:ctk.CTkBaseClass) -> None:
        """ Remove widget from frame """
        this_widget.pack_forget()
        
    @staticmethod
    def setOptions(new_values:list[str] | tuple[str, ...], optionMenu:ctk.CTkOptionMenu, new_variable:ctk.StringVar):
        """ Changes values of optionMenu and  set the corresponding variable to the options """
        optionMenu.configure(values=new_values)
        new_variable.set(new_values[0])
        
    @staticmethod
    def changeOptions(options:list[str], OptionList:ctk.CTkOptionMenu) -> None:
        """ Updates optionmenu with options """
        OptionList.configure(values=options)
        
    @staticmethod
    def changeCommand(widgets:ctk.CTkBaseClass, new_command: Callable[..., Any]) -> None:
        """ Changes command of a widget """
        widgets.configure(command=new_command)
        
    @staticmethod
    def changeText(widgets:ctk.CTkLabel, text:str) -> None:
        """ Changes text of widgets """
        widgets.configure(text=text)

    @staticmethod
    def place(widget:ctk.CTkBaseClass,padx:int = 10, pady:int = 10) -> None:
        """ Places widget after the last widget present on frame """
        widget.pack(padx=padx,pady=pady)
        
    @staticmethod
    def clear_entry(widget: ctk.CTkEntry) -> None:
        """ Clears entry widget """
        widget.delete(0, tk.END)

class Decryption:
    """ Handles file decryption  """
    
    
    @staticmethod
    def is_encrypted(file:io.BufferedReader) -> bool:
        """ Checks if file is encrypted """
        return msoffcrypto.OfficeFile(file).is_encrypted()
    
    @staticmethod
    def decrypting_file(file_path:Path, decrypted:io.BytesIO, password:str) -> tuple[bool, io.BytesIO]:
        """ Decrypts the file and extracts content into BytesIO """
        
        success = False
        try:
            with open(file_path.resolve(), "rb") as f:
                file = msoffcrypto.OfficeFile(f)
                file.load_key(password=password)  
                file.decrypt(decrypted)
                success = True
                
        except msoffcrypto.exceptions.InvalidKeyError as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))

        except msoffcrypto.exceptions.DecryptionError as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
            
        except Exception as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
            
        return (success, decrypted)
    
    @staticmethod
    def is_encrypted_wrapper(queue: Queue, file_path:Path) -> None:
        """ Wrapper for is_encrypted  """
        result = None

        try:
            if(file_path.exists()):
                with open(file_path.resolve(), 'rb') as f:
                    result = Decryption.is_encrypted(f)
                    
        except Exception as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
            result = None

        queue.put(result)
    
    @staticmethod
    def fetch_decrypted_file(queue: Queue, file_path: Path, skip:int = 0) -> None:
        """ Fetches file that is not encrypted """
        
        result: list[Optional[list[str]] | Optional[dict[str, pd.DataFrame]]] = [None, None]
        
        try:
            data = pd.read_excel(io=file_path.resolve(), sheet_name=None, skiprows=skip)
            
            if(data):
                sheets = list(data.keys())
                
                for i in sheets: dataRefine(data[i])
                
                result[0] = sheets
                result[1] = data
            
        except Exception as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
        
        queue.put(tuple(result))
    
    @staticmethod
    def fetch_encrypted_file(queue: Queue, file_path:Path, password:str, skip:int = 0) -> None:
        """ Fetch content from encrypted file """
        
        result: list[Optional[list[str]] | Optional[dict[str, pd.DataFrame]]] = [None,None]

        try:
            with io.BytesIO() as decrypted:
                success, file = Decryption.decrypting_file(file_path, decrypted, password)
            
                if(success):
                    
                    data = pd.read_excel(io=file,sheet_name=None,skiprows=skip, dtype=str)
                    sheets = list(data.keys())
                    
                    if(data):
                        for i in sheets: dataRefine(data[i])
                        
                        result[0] = sheets
                        result[1] = data
                        
        except Exception as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
        
        queue.put(tuple(result))    

class MailingWrapper:
    """ Wrapper for Mailing class """
    
    month = 'jan'
    year = 2024
        
    def change_state(self, month:MonthList, year:int) -> 'MailingWrapper':
        """ Set month and year of this class """
        self.month = month
        self.year = year
        return self
    
    def attempt_mail_process(self, pdf_path:Path, id:str, toAddr:str, queue:Queue) -> None:
        """ Updates queue when mailing process is done """
        
        ok = (
                Mailing(**MAIL_CRED, error_log=ERROR_LOG).login()
                .addTxtMsg(f"Please find attached below the salary slip of {self.month.capitalize()}-{self.year}",'plain')
                .addAttach(pdf_path,f'employee_{file_clean(id)}.pdf')
                .addDetails(f"Salary slip of {self.month.capitalize()}-{self.year}")
                .sendMail(toAddr).resetMIME().destroy().status
            )
        
        queue.put(ok)
        
            
    async def _sendMail(self, mail:'AsyncMailing', pdf_path: Path, id:str, toAddr:str) -> bool:
        """ Wrapper for continuous mail sending """
        
        def __make_msg__(mailing: MailingWrapper) -> MIMEMultipart:
            msg = AsyncMessage(ERROR_LOG)
            msg.addTxtMsg(f"Please find attached below the salary slip of {mailing.month.capitalize()}-{mailing.year}",'plain')
            msg.addAttach(pdf_path, f'employee_{id}.pdf')
            msg.addDetails(f"Salary slip of {mailing.month.capitalize()}-{mailing.year}")
            return msg.get_MIME()
            
        msg = await to_thread(__make_msg__, self)
        await mail.sendMail(toAddr, msg)
        return mail.status
        
    @staticmethod
    async def report(tasks:list[types.CoroutineType]) -> int:
        result:list[bool] = await gather(*tasks)
        
        return len(list(filter(lambda x: x, result)))
    
    def massMail(self, data:pd.DataFrame, code_column:str, email_col:str, dir_path:Path, queue:Queue) -> None:
        """ Sends email on basis of pdf files present in chosen directory """
                
        count,total = 0, 0
        email_server = AsyncMailing(**MAIL_CRED,error_log=ERROR_LOG)
        columns = {j:i for i,j in enumerate(data.columns)}
        
        async def __dont_know__():
            nonlocal email_server, columns, total, count
            mails = []
            
            try:
                email_server = await email_server.login()
                
                for _, row in data.iterrows():
                    
                    try:
                        emp_id = row.values[columns[code_column]]
                        file_name = f"employee_{file_clean(emp_id)}.pdf"
                        
                        pdf_path = dir_path.joinpath(file_name)
                        
                        if not pdf_path.exists():
                            continue
                        
                        total+=1
                        toAddr = row.values[columns[email_col]]
                        mails.append(self._sendMail(email_server, deepcopy(pdf_path), deepcopy(emp_id), deepcopy(toAddr)))

                    except Exception as e:
                        ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
                        
                count = await MailingWrapper.report(mails)
                queue.put((count, total))
                    
            except Exception as e:
                ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
            
            await email_server.destroy()
        
        run(__dont_know__())
        
class DatabaseWrapper:
    """ Wrapper for DataBase """
    DATABASE = Database(ERROR_LOG)
    
    def connectToDatabase(self) -> Database:
        """ Connect to database """
        return self.DATABASE.connectDatabase(self.host,self.user,self.password,self.database)
    
    def __init__(self, host:str, user:str, password:str, database:str):
        self.host = host
        self.user = user
        self.password = password
        self.database = database
    
    def check_table(self, queue:Queue) -> None:
        """ Wrapper for showTables """
        queue.put(self.connectToDatabase().showTables())
        self.endThis()
    
    def get_data(self, queue: Queue, institute: InstituteList, type: TypeList, year:int, month: MonthList) -> None:
        """ Wrapper for fetchAll """
        queue.put(self.connectToDatabase().fetchAll(month,year,institute,type))
        self.endThis()
        
    def create_table(self, queue: Queue, institute: InstituteList, type: TypeList, year:int, month: MonthList, data_columns: list[str]) -> None:
        """ Wrapper for create table """
        create_result = self.connectToDatabase().createData(month,year,data_columns,institute,type)
        queue.put(create_result)
        self.endThis()
        
    def fill_table(self, queue: Queue, institute: InstituteList, type: TypeList, year:int, month: MonthList, data: pd.DataFrame):
        """ Attempts to insert data or update data in db (Doesn't ask for updation)"""
        upsert_result = self.connectToDatabase().updateData(data,month,year,institute,type)
        queue.put(upsert_result)
        self.endThis()
        
    def delete_table(self, queue: Queue, institute: InstituteList, type: TypeList, year:int, month: MonthList):
        """ Delete Table """
        delete_result = self.connectToDatabase().dropTable(month, year, institute, type)
        queue.put(delete_result)
        self.endThis()
        
    def endThis(self):
        """  End connection """
        self.DATABASE.endDatabase()
        
class PandaWrapper:
    """ Panda wrapper for process """
    def __init__(self, servitor:pd.DataFrame, identification_rosette:str, pdf_template:PDFTemplate) -> None:
        self.servitor = servitor
        self.identification_rosette = identification_rosette
        self.pdf = pdf_template
        self.columns = {j:i for i,j in enumerate(servitor.columns)}
        self.vars: dict[str, NullStr] = {}
        self.html_file: str = ""
        self.column_auspex: dict[str, NullStr] = {}
    
    def load_scriptures(self) -> None:
        
        if(self.pdf.chosen_json is not None): 
            self.column_auspex = {i:mapping(list(self.columns.keys()), j) for i,j in self.pdf.load_json(self.pdf.chosen_json).items()}
            self.column_auspex[CODE_COL] = self.identification_rosette
            self.column_auspex["branch"] = "Sion"
    
    def find_by_id(self, queue: Queue,emp_id:str):
        """ Finds employee by emp_id in identification_rosette column of Dataframe """
        search_result = self.servitor[self.servitor[self.identification_rosette]==emp_id]
        if(not search_result.empty):
            queue.put(self.litany_of_auspex(search_result.iloc[[0]]))
        else:
            queue.put(self.litany_of_failure())
            
    def litany_of_failure(self): 
        """ The labyrinth does not bears the knowledge you wish to seek, Tech Adept """
        return None
    
    def litany_of_auspex(self, search_result:pd.DataFrame):
        """ Praised be the machine spirit of the blessed augurs (Returns data about id, which should exist)"""
        return search_result
    
    @staticmethod
    async def get_success(tasks:list[types.CoroutineType]) -> tuple[int,int]:
        result = await gather(*tasks)
        
        success, total = 0, 0
        
        for i in result:
            status, _ = i
            if(status): success+=1
            total += 1
        
        return success, total
    
    def litany_of_scrolls(self, queue: Queue, dropzone:Path, month:MonthList, year:int):
        """ Oh holy generator, grants us the sacred ink of thou's blessed blood (BulkPrint Process)"""
        task_completed = 0
        total_task = self.servitor.shape[0]
        tasks = []
        self.load_scriptures()
        
        for data in self.servitor.itertuples(index=False):
            emp_data = {key: val or "-" if (val not in set(self.columns.keys())) else text_clean(data[self.columns[val]]) for key,val in self.column_auspex.items()}
            emp_data.update({'month':month.capitalize(),'year': str(year)})
            
            if(self.pdf.chosen_html is not None):
                html_content = self.pdf.render_html(self.pdf.chosen_html, emp_data)
                    
                tasks.append(PDFGenerator.generate_many_pdf(emp_data.get(CODE_COL,"none"), html_content,dropzone))
            
        task_completed, _ = run(PandaWrapper.get_success(tasks)) # type: ignore

        queue.put((task_completed, total_task))
        
    def litany_of_scroll(self, queue: Queue, data_slate_path:Path, month:str, year:int, emp_id:str):
        """ Oh holy generator, grants us the sacred ink of thou's blessed blood (SinglePrint Process)"""
        
        search_result = self.servitor[self.servitor[self.identification_rosette]==emp_id]
        self.load_scriptures()
        
        
        if(not search_result.empty):
            data = search_result.iloc[0].to_numpy()
            emp_data = {key: val or "-" if (val not in set(self.columns.keys())) else text_clean(data[self.columns[val]]) for key,val in self.column_auspex.items()}
            emp_data.update({'month':month.capitalize(),'year': str(year)})
            
            if self.pdf.chosen_html is not None:
                html_content = self.pdf.render_html(self.pdf.chosen_html, emp_data)
                        
                status, msg = PDFGenerator.generate_one_pdf(emp_data.get(CODE_COL,"none"), html_content, data_slate_path)
            else:
                status, msg = None, "Something went wrong"
            queue.put((status,msg))
            
        else:
            queue.put((None, f"Employee Code '{emp_id}' was not found"))

class TemplateGenerator:
    
    counter: int = 0
    
    @staticmethod
    def make_everything(elems: dict[str, str]) -> tuple[str, dict[int, str]]:
        html_string = ""
        jsonDict: dict[int, str] = {}
        
        for key, val in elems.items():
            count = TemplateGenerator.counter
            html_string += "<tr><td>{0}:</td>\n<td><span>{{{{{1}}}}}</span></td></tr>\n".format(key, count) # { escapes {
            jsonDict[count] = text_clean(val)
            TemplateGenerator.counter += 1
            
        return html_string, jsonDict

    @staticmethod
    def make_template(file_name: str, data: dict[str, pd.DataFrame], queue: Queue) -> None:
        
        TemplateGenerator.counter = 0
        memo: dict[str, str] = {}
        jsonDict: dict[int, str] = {}
        
        for sheet in TEMPLATE_SHEET:
            sheetName = sheet.strip().replace(" ","_")
            memo[sheetName] = ""
            
            sheetData = data.get(sheet, pd.DataFrame())
            
            if(not all([ (i in sheetData.columns) for i in TEMPLATE_COLUMN])):
                continue
            
            _name, _title = TEMPLATE_COLUMN
            columns = sheetData.columns
            column_memo = {j: i for i,j in enumerate(columns)}
            
            if(((name := mapping(columns, _name)) is not None) and ((title := mapping(columns, _title)) is not None)):
                row_data = sheetData.to_numpy()
                
                memo[sheetName], tempDict = TemplateGenerator.make_everything({row[column_memo[name]]: row[column_memo[title]] for row in row_data})
            
                jsonDict.update(tempDict)
            
        html_string = TEMPLATE % memo
        file_name = file_name.replace(" ", "_")
        status, msg = PDF_TEMPLATE._load_defaults(PDF_TEMPLATE.html_path, 'html', **{file_name: html_string})
        queue.put((status, msg))
        status, msg = PDF_TEMPLATE._load_defaults(PDF_TEMPLATE.json_path, 'json', **{file_name: json.dumps(jsonDict)})
        queue.put((status, msg))
    
    @staticmethod
    def make_excel(file_name: Path, data: dict[str, pd.DataFrame], queue: Queue) -> None:
        
        TemplateGenerator.counter = 0
        columnsList: list[str] = []
        
        for sheet in TEMPLATE_SHEET:

            sheetData = data.get(sheet, pd.DataFrame())
            if(not all([ (i in sheetData.columns) for i in TEMPLATE_COLUMN])):
                continue
            
            _name, _title = TEMPLATE_COLUMN
            columns = sheetData.columns
            column_memo = {j: i for i,j in enumerate(columns)}
            
            if(((_ := mapping(columns, _name)) is not None) and ((title := mapping(columns, _title)) is not None)):
                row_data = sheetData.to_numpy()
                
                for row in row_data:
                    columnsList.append(row[column_memo[title]])
                
        
        try:
            a = pd.DataFrame({}, columns=columnsList)
            a.to_excel(file_name, index=False)
            
            queue.put((True, "Excel File generated successfully"))
        except Exception as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
            
            queue.put((False, "Excel File generation failed"))

class MailCover(BaseTemplate):
    """ Mailing Cover to choose either single-mailing / bulk-mailing """
    
    def __init__(self,outer:App):
        super().__init__(outer)
        
        ctk.CTkLabel(master=self.frame , text="Mailing Option", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)
        ctk.CTkLabel(master=self.frame , text="Select mailing option", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10)
        ctk.CTkButton(master=self.frame, text='Single Mail', command=self.single, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10, padx=10)
        ctk.CTkButton(master=self.frame, text='Bulk Mail', command=self.many, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10, padx=10)
        ctk.CTkButton(master=self.frame, text='Back', command=self.back_to_landing, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10, padx=10)

    def single(self) -> None:
        """ Redirect to Single Mailing """
        self.switch_screen(SendMail)
    
    def many(self) -> None:
        """ Redirect to Bulk Mailing """
        self.switch_screen(SendBulkMail)

    def back_to_landing(self) -> None:
        """ Returns back to Interface """
        self.clear_data()
        self.clear_queue()
        self.switch_screen(Interface)
        
class SendMail(BaseTemplate):
    """ Page for single mail """    
    chosen_institute = ctk.StringVar(value='Somaiya')
    chosen_type = ctk.StringVar(value='Teaching')
    chosen_month = ctk.StringVar(value='Jan')   
    
    def __init__(self,outer:App):
        super().__init__(outer)
        
        ctk.CTkLabel(master=self.frame , text="Single Mail", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame , text="Selected PDF File:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.file = ctk.CTkEntry(master=frame , width=100, text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"))
        self.file.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        self.browse_button = ctk.CTkButton(master=self.frame , text="Browse", command=self.browse_file, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.browse_button.pack(pady=10)
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Institute:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.toggle_institute = ctk.CTkOptionMenu(master=frame,variable=self.chosen_institute,values=list(self.outer.TOGGLE.keys()),command=self.changeType,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.toggle_institute.pack(pady=10, padx=10)
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Type:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.toggle_type = ctk.CTkOptionMenu(master=frame,variable=self.chosen_type,values=self.outer.TOGGLE[list(self.outer.TOGGLE.keys())[0]],button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.toggle_type.pack(pady=10, padx=10)
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Month:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.toggle_month = ctk.CTkOptionMenu(master=frame,variable=self.chosen_month,values=[str(i).capitalize() for i in self.outer.MONTH],button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.toggle_month.pack(pady=10, padx=10)
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame , text="Enter Year:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(side='left',pady=10, padx=10)
        self.year = ctk.CTkEntry(master=frame ,placeholder_text="Eg. 2024", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=100)
        self.year.pack(side='left',pady=10, padx=10)
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame , text="Enter Employee ID:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(side='left',pady=10, padx=10)
        self.emp_id = ctk.CTkEntry(master=frame ,placeholder_text="Eg. 2200317", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=150)
        self.emp_id.pack(side='left',pady=10, padx=10)
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame , text="Enter Employee Email:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(side='left',pady=10, padx=10)
        self.email = ctk.CTkEntry(master=frame ,placeholder_text="Eg. sample@somaiya.edu", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.email.pack(side='left',pady=10, padx=10)
        frame.pack()
        
        self.send_button = ctk.CTkButton(master=self.frame , text='Send Mail', command=self.send_mail, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.send_button.pack(pady=10,padx=10)
        
        self.back_button = ctk.CTkButton(master=self.frame , text='Back', command=self.back, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back_button.pack(pady=10,padx=10)
        
        self.quit = ctk.CTkButton(master=self.frame,text='Quit the process',command=self.cancel_thread_wrapper,fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        
        self.to_disable = list(self.get_widgets_to_disable())
        
    def cancel_thread_wrapper(self) -> None:
        """  Wrapper for cancelling thread and process """
        self.cancel_thread()
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        tkmb.showinfo('Email Status','Email Sending Stopped')
    
    def send_mail_thread_wrapper(self) -> None:
        """ Handles gui during mailing process """
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.send_button,self.quit)
        done:bool | None = None
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                done = self.QUEUE.get()
                self.clear_queue()
                break
        
        if(done):
            tkmb.showinfo('Email Status',"Email Send Successfully")
        else:
            tkmb.showinfo('Email Status','Email was not send successfully')
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
        
    def changeType(self, event:Optional[tk.Event] = None):
        """ changes type according to institute """
        institute = self.chosen_institute.get()
        type = self.outer.TOGGLE[institute]
        GUI_Handler.setOptions(type,self.toggle_type,self.chosen_type)

    def send_mail(self) -> None:
        """ Sends mail """
        file_path = self.file.get()
        toAddr = text_clean(self.email.get())
        month = self.chosen_month.get()
        year = text_clean(self.year.get())
        id = text_clean(self.emp_id.get())
        
        if(not year_check(year)):
            tkmb.showwarning('Year Check','Improper Year format')
            return
        
        if(not email_check(toAddr)):
            tkmb.showwarning('Email Check',"Improper Email Address format")
            return
        
        if(not (re.match(r'^(\w+|\d+)\Z',id))):
            tkmb.showwarning('Employee ID check','Improper Employee ID format')
            return
        
        if(not os.path.isfile(file_path)):
            tkmb.showwarning('File Check','Selected path is not a file')
            return
        
        if(self.can_start_thread()):
            self.process = Process(target=MailingWrapper().change_state(month,year).attempt_mail_process,kwargs={'pdf_path': file_path,'toAddr': toAddr,'id': id,'queue':self.QUEUE},daemon=True)
            self.thread = Thread(target=self.send_mail_thread_wrapper,daemon=True)
            self.thread.start()
            self.process.start()
            
        else:
            tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")

    def browse_file(self) -> None:
        """ Browse for file to send """
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", ".pdf")],initialdir=Path(APP_PATH).parent)
        
        if file_path:
            GUI_Handler.change_file_holder(self.file,file_path)
            
    def back(self) -> None:
        """ Back to MailCover """
        self.hide()
        self.clear_data()
        self.outer.CHILD[MailCover.__name__].appear()

class SendBulkMail(BaseTemplate):
    """ Page for Bulk Mailing """
    chosen_institute = ctk.StringVar(value='Somaiya')
    chosen_type = ctk.StringVar(value='Teaching')
    chosen_month = ctk.StringVar(value='Jan')  
    email_server = Mailing(**MAIL_CRED,error_log=ERROR_LOG)
    count, total = 0, 0
    
    def __init__(self,outer:App):
        super().__init__(outer)

        ctk.CTkLabel(master=self.frame , text="Bulk Mail", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)

        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame , text="Selected Folder Location:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.folder = ctk.CTkEntry(master=frame , width=100, text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"))
        self.folder.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        self.browse_button = ctk.CTkButton(master=self.frame , text='Browse', command=self.folder_browse, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.browse_button.pack(pady=10)
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Institute:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_institute =ctk.CTkOptionMenu(master=frame,variable=self.chosen_institute,values=list(self.outer.TOGGLE.keys()),command=self.changeType,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.entry_institute.pack(pady=10, padx=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Type:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.toggle_type = ctk.CTkOptionMenu(master=frame,variable=self.chosen_type,values=self.outer.TOGGLE[list(self.outer.TOGGLE.keys())[0]],button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.toggle_type.pack(pady=10, padx=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Month:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_month = ctk.CTkOptionMenu(master=frame,variable=self.chosen_month,values=[str(i).capitalize() for i in self.outer.MONTH],button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.entry_month.pack(pady=10, padx=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame , text="Enter Year:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(side='left',pady=10, padx=10)
        self.entry_year = ctk.CTkEntry(master=frame ,placeholder_text="Eg. 2024", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=100)
        self.entry_year.pack(side='left',pady=10, padx=10)
        frame.pack()
        
        self.exists_table = ctk.CTkButton(master=self.frame , text='Check Database', command=self.table_exists, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.exists_table.pack(pady=10, padx=10)
        
        self.mail_button = ctk.CTkButton(master=self.frame , text='Send Emails', command=self.send_mail, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
    
        self.back = ctk.CTkButton(master=self.frame , text='Back', command=self.button_mailing, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back.pack(pady=10, padx=10)
        
        self.quit = ctk.CTkButton(master=self.frame,text='Quit the process',command=self.cancel_thread_wrapper,fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        
        self.to_disable = list(self.get_widgets_to_disable())
        
        
    def changeType(self, event:Optional[tk.Event] = None):
        """ changes type according to institute """
        institute = self.chosen_institute.get()
        type = self.outer.TOGGLE[institute]
        GUI_Handler.setOptions(type,self.toggle_type,self.chosen_type)
        
    def cancel_thread_wrapper(self) -> None:
        """ Wrapper for cancelling thread and process """
        self.cancel_thread()
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        tkmb.showinfo('Email Status','Email Sending Stopped')
        self.email_result(self.count,self.total)
            
    def send_mail_thread_wrapper(self) -> None:
        """ You know the drill """
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.mail_button,self.quit)
        GUI_Handler.changeCommand(self.quit,self.cancel_thread_wrapper)
        
        self.count, self.total = 0, 0
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                curr = self.QUEUE.get()
            
                if(curr=="Done"):
                    self.clear_queue()
                    break
                
                elif(isinstance(curr,tuple) and len(curr)==2):
                    self.count, self.total = curr
                
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        self.email_result(self.count,self.total)
    
    
    def email_result(self, count:int, total:int) -> None:
        """ Shows result after mailing is complete """
        
        if((not self.stop_flag) and self.count>0):
            tkmb.showinfo('Email Status',f"{count}/{total} emails were send successfully")
        elif((not self.stop_flag) and self.total>0):
            tkmb.showinfo('Email Status','No emails were send')

    def send_mail(self) -> None:
        """ See name """
        file_path = self.folder.get()
        month = self.chosen_month.get()
        year = self.entry_year.get()
        institute = self.chosen_institute.get()
        type = self.chosen_type.get()
        
        if(BaseTemplate.data is None):
            tkmb.showwarning('Data Status',f'Data for {institute.capitalize()} {type.capitalize()} {month.capitalize()}/{year} was not found. Please do the necessary')
            return
        
        code_col = mapping(BaseTemplate.data.columns,'HR EMP CODE')
        email_col = None
        all_columns = BaseTemplate.data.columns
        
        for i in all_columns:
            if ("mail" in i.lower()):
                email_col = i
                break

        if(code_col is None or email_col is None):
            tkmb.showwarning('Column Missing','HR EMP CODE or Mail Column was not found')
            return
    
        if(not year_check(year)):
            tkmb.showwarning('Year Check','Improper Year format')
            return
                
        if(not file_path):
            tkmb.showwarning('File Path Check','File Path is empty')
            return
        
        if(not os.path.isdir(file_path)):
            tkmb.showwarning('File Check','Selected path is not a directory')
            return
                
        if(self.can_start_thread()):
            self.process = Process(target=MailingWrapper().change_state(month,year).massMail,kwargs={'data': BaseTemplate.data,'code_column': code_col,'email_col': email_col,'dir_path': file_path,'queue':self.QUEUE},daemon=True)
            self.thread = Thread(target=self.send_mail_thread_wrapper,daemon=True)
            
            self.thread.start()
            self.process.start()
        
        else:
            tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")

    def folder_browse(self) -> None:
        current_dir = os.path.dirname(APP_PATH)
        entry_folder = filedialog.askdirectory(initialdir=current_dir)

        if entry_folder:
            GUI_Handler.change_file_holder(self.folder,entry_folder)
            
    def table_exists_thread(self):
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.changeCommand(self.quit,self.table_exists_cancel_thread)
        GUI_Handler.place_after(self.exists_table,self.quit)
        
        data: Optional[pd.DataFrame] = None
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                data = self.QUEUE.get()
                self.clear_queue()
                break
            
        if(data is not None):
            GUI_Handler.place_after(self.exists_table,self.mail_button)
            tkmb.showinfo('Database Status',f"Table {self.chosen_institute.get()}_{self.chosen_type.get()}_{self.chosen_month.get()}_{self.entry_year.get()} exists in database. Mass Mailing available")
            BaseTemplate.data = data
        else:
            tkmb.showinfo('Database Status',f"Table {self.chosen_institute.get()}_{self.chosen_type.get()}_{self.chosen_month.get()}_{self.entry_year.get()} doesn't exist in database. Please upload data to database")
            
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
    def table_exists_cancel_thread(self):
        self.cancel_thread()
        
        tkmb.showinfo('Database Status',"Process Halted")
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
            
    
    def table_exists(self):
        month = self.chosen_month.get()
        year = self.entry_year.get()
        
        if(not year_check(year)):
            tkmb.showwarning('Year Check','Improper Year format')
            return
                
        if(self.can_start_thread()):
            self.process = Process(target=DatabaseWrapper(**self.outer.CRED).get_data,kwargs={'queue':self.QUEUE,'institute': self.chosen_institute.get(),'type': self.chosen_type.get(),'year': year,'month': month},daemon=True)
            self.thread = Thread(target=self.table_exists_thread,daemon=True)
            
            self.thread.start()
            self.process.start()

        else:
            tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")

    def button_mailing(self) -> None:
        self.clear_data()
        GUI_Handler.remove_widget(self.mail_button)
        self.switch_screen(MailCover)

class Login(BaseTemplate):
    """ Login Page """
    
    def __init__(self,outer:App) -> None:
        super().__init__(outer)

        ctk.CTkLabel(master=self.frame , text="Admin Login Page", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame,text="Username:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.user_entry = ctk.CTkEntry(master=frame,placeholder_text="Username" , width=250, text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"))
        self.user_entry.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame , text="Password:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.user_password = ctk.CTkEntry(master=frame,placeholder_text="Password" ,show="*",width=250, text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"))
        self.user_password.pack(padx=10,pady=10,side='left')
        frame.pack()

        ctk.CTkButton(master=self.frame, text='Login', command=self.login, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10, padx=10)
        ctk.CTkButton(master=self.frame, text='Exit', command=self.quit, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10, padx=10)

    def login(self) -> None:
        """ Check username and password """
        known_user = 'admin' # actual username
        known_pass = 'kjs2024' # actual password
        
        username = self.user_entry.get() if not IS_DEBUG else known_user
        password = self.user_password.get() if not IS_DEBUG else known_pass
        
        if(not username or not password):
            tkmb.showwarning(title='Empty Field', message='Please fill the all fields')
            return 
        
        if known_user == username and known_pass == password:
            tkmb.showinfo(title="Login Successful", message="You have logged in Successfully")
            ERROR_LOG.write_info('User Logged in')
            self.switch_screen(MySQLLogin)
        else:
            tkmb.showwarning(title='Wrong password', message='Please check your username/password')

    def quit(self):
        if messagebox.askyesnocancel("Confirmation", f"Are you sure you want to exit"):
            self.outer.APP.quit()

class MySQLLogin(BaseTemplate):
    def __init__(self,outer:App) -> None:
        super().__init__(outer)

        ctk.CTkLabel(master=self.frame , text="MySQL Login Page", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)

        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame,text="Host:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.host = ctk.CTkEntry(master=frame ,placeholder_text="Host", width=250, text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"))
        self.host.insert(0,"localhost")
        self.host.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame,text="Username:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.user = ctk.CTkEntry(placeholder_text="Username",master=frame , width=250, text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"))
        self.user.insert(0,"root")
        self.user.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame,text="Password:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.password = ctk.CTkEntry(master=frame,placeholder_text="Password" , show="*",width=250, text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"))
        self.password.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame,text="Database:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.database = ctk.CTkEntry(master=frame, placeholder_text="Database" , width=250, text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"))
        self.database.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        ctk.CTkButton(master=self.frame, text='Continue', command=self.next, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10, padx=10)
        ctk.CTkButton(master=self.frame, text='Back', command=self.back, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10, padx=10)

    # back to login page
    def back(self):
        self.hide()
        self.outer.CHILD[Login.__name__].appear()
            
    # moves to next page (landing) is mysql connection was established
    def next(self):
        host = self.outer.CRED['host'] = self.host.get() if not IS_DEBUG else MYSQL_CRED['host']
        user = self.outer.CRED['user'] = self.user.get() if not IS_DEBUG else MYSQL_CRED['user']
        password = self.outer.CRED['password'] = self.password.get() if not IS_DEBUG else MYSQL_CRED['password']
        database = self.outer.CRED['database'] = self.database.get() if not IS_DEBUG else MYSQL_CRED['database']
        
        if(host and user and password and database):
        
            if(self.outer.DB.connectDatabase(**self.outer.CRED).isConnected()):
                self.outer.DB.endDatabase()
                tkmb.showinfo('MySQL Status','MySQL Connection Established')
                self.switch_screen(Interface)
            else:
                tkmb.showinfo('MySQL Status','MySQL Connection Failed')
        
        else:
            tkmb.showwarning(title='Empty Field', message='Please fill the all fields')

class Interface(BaseTemplate):
    def __init__(self, outer:App) -> None:
        super().__init__(outer)
        
        ctk.CTkLabel(master=self.frame , text="Excel-To-PDF Generator", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)
        
        self.preview_db = ctk.CTkButton(master=self.frame , text='Preview Existing Data', command=self.preview, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.preview_db.pack(pady=10, padx=10)
        
        self.upload_button = ctk.CTkButton(master=self.frame , text='Upload Excel data', command=self.upload, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.upload_button.pack(pady=10, padx=10)
        
        self.mail_button = ctk.CTkButton(master=self.frame , text='Send Mail', command=self.mail, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.mail_button.pack(pady=10, padx=10)
        
        self.delete_button = ctk.CTkButton(master=self.frame , text='Delete Tables', command=self.delete, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.delete_button.pack(pady=10, padx=10)
        
        self.template_button = ctk.CTkButton(master=self.frame , text='Generate Templates', command=self.template, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.template_button.pack(pady=10, padx=10)
        
        self.template_download_button = ctk.CTkButton(master=self.frame , text='Download Templates', command=self.template_download, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.template_download_button.pack(pady=10, padx=10)
        
        self.back = ctk.CTkButton(master=self.frame , text='Back', command=self.back_to_mysql, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back.pack(pady=10, padx=10)
        
        self.quit = ctk.CTkButton(master=self.frame,text='Quit the process',command=self.cancel_thread_wrapper,fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.to_disable = list(self.get_widgets_to_disable())
        
    def back_to_mysql(self) -> None:
        """ back to mysql setup page """
        self.switch_screen(MySQLLogin)

    def upload(self) -> None:
        """ proceeds to upload page """
        self.switch_screen(FileInput)

    def mail(self) -> None:
        self.switch_screen(MailCover)
    
    def check_database_thread(self):
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.preview_db,self.quit)
        table = {}
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                table = self.QUEUE.get()
                self.clear_queue()
                break
        
        if(table):
            self.outer.CHILD[DataPreview.__name__].tables = table
            self.outer.CHILD[DataPreview.__name__].changeData()
            self.switch_screen(DataPreview)
        else:
            tkmb.showerror("MySQL Error","No data available to preview")

        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
    
    def cancel_thread_wrapper(self) -> None:
        self.cancel_thread()
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        tkmb.showinfo('Fetch Status','Data Fetching Stopped')

    # proceeds to pre-existing data page
    def preview(self) -> None:
        
        if(self.can_start_thread()):
            self.process = Process(target=DatabaseWrapper(**self.outer.CRED).check_table,kwargs={'queue':self.QUEUE},daemon=True)
            self.thread = Thread(target=self.check_database_thread,daemon=True)

            self.thread.start()
            self.process.start()
            
        else:
            tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
    
    
    def delete_database_thread(self):
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.delete_button,self.quit)
        table = {}
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                table = self.QUEUE.get()
                self.clear_queue()
                break
        
        if(table):
            self.outer.CHILD[DataPeek.__name__].tables = table
            self.outer.CHILD[DataPeek.__name__].changeData()
            self.switch_screen(DataPeek)
        else:
            tkmb.showerror("MySQL Error","No data available to delete")
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
    
    def delete(self) -> None:
        
        if(self.can_start_thread()):
            self.process = Process(target=DatabaseWrapper(**self.outer.CRED).check_table,kwargs={'queue':self.QUEUE},daemon=True)
            self.thread = Thread(target=self.delete_database_thread,daemon=True)

            self.thread.start()
            self.process.start()
            
        else:
            tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
            
    def template(self) -> None:
        self.switch_screen(TemplateInput)
        
    def template_download(self) -> None:
        self.switch_screen(TemplateGeneration)

class DataPreview(BaseTemplate):
    """ Previews data in database """
    
    entry_year = ctk.StringVar(value='2024')
    entry_month = ctk.StringVar(value='Jan')
    entry_institute = ctk.StringVar(value='Somaiya')
    entry_type = ctk.StringVar(value='Teaching')
    tables:dict[str,dict[str,dict[str,set[str]]]] = {}
    
    def __init__(self,outer:App) -> None:
        super().__init__(outer)

        ctk.CTkLabel(master=self.frame , text="Data Preview", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Institute:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_instituteList = ctk.CTkOptionMenu(master=frame,variable=self.entry_institute,values=[],command=self.changeInstitute,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.entry_instituteList.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Type:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_typeList = ctk.CTkOptionMenu(master=frame,variable=self.entry_type,values=[],command=self.changeType,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.entry_typeList.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Year:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_yearList = ctk.CTkOptionMenu(master=frame,variable=self.entry_year,values=[],command=self.changeYear,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.entry_yearList.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Month:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_monthList = ctk.CTkOptionMenu(master=frame,variable=self.entry_month,values=[],button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.entry_monthList.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        self.fetch_data = ctk.CTkButton(master=self.frame , text='Continue', command=self.getData, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.fetch_data.pack(pady=10, padx=10)
        
        self.back = ctk.CTkButton(master=self.frame , text='Back', command=self.back_to_interface, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back.pack(pady=10, padx=10)
        
        self.quit = ctk.CTkButton(master=self.frame,text='Quit the process',command=self.cancel_thread_wrapper,fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.to_disable = list(self.get_widgets_to_disable())
        
    def back_to_interface(self) -> None:
        """ Moves back to the earlier page """
        self.tables = {}
        self.switch_screen(Interface)
            
    def changeData(self):
        """ Initialization of the optionmenus """     
        curr_institute = list(self.tables.keys())
        curr_type = list(self.tables[curr_institute[0]].keys())
        curr_year = list(self.tables[curr_institute[0]][curr_type[0]].keys())
        curr_month = sorted(list(self.tables[curr_institute[0]][curr_type[0]][curr_year[0]]),key=lambda x:self.outer.MONTH[x])
        GUI_Handler.setOptions(curr_type,self.entry_typeList,self.entry_type)
        GUI_Handler.setOptions(curr_institute,self.entry_instituteList,self.entry_institute)
        GUI_Handler.setOptions(curr_year,self.entry_yearList,self.entry_year)
        GUI_Handler.setOptions(curr_month,self.entry_monthList,self.entry_month)
        
    def changeInstitute(self,event:tk.Event) -> None:
        """ Changes Institute Type """
        curr_institute = self.entry_institute.get()
        curr_type = list(self.tables[curr_institute].keys())
        curr_year = list(self.tables[curr_institute][curr_type[0]].keys())
        curr_month = sorted(list(self.tables[curr_institute][curr_type[0]][curr_year[0]]),key=lambda x:self.outer.MONTH[x])

        GUI_Handler.setOptions(curr_type,self.entry_typeList,self.entry_type)
        GUI_Handler.setOptions(curr_year,self.entry_yearList,self.entry_year)
        GUI_Handler.setOptions(curr_month,self.entry_monthList,self.entry_month)
        
    def changeType(self,event:tk.Event):
        """ Changes Type Type """
        curr_institute = self.entry_institute.get()
        curr_type = self.entry_type.get()
        curr_year = list(self.tables[curr_institute][curr_type].keys())
        curr_month = sorted(list(self.tables[curr_institute][curr_type][curr_year[0]]),key=lambda x:self.outer.MONTH[x])
        
        GUI_Handler.setOptions(curr_year,self.entry_yearList,self.entry_year)
        GUI_Handler.setOptions(curr_month,self.entry_monthList,self.entry_month)
        
    # updates menu when year changes
    def changeYear(self,event:tk.Event) -> None:
        """ Changes the year """
        curr_institute = self.entry_institute.get()
        curr_type = self.entry_type.get()
        curr_year = self.entry_year.get()
        curr_month = sorted(list(self.tables[curr_institute][curr_type][curr_year]),key=lambda x:self.outer.MONTH[x])

        GUI_Handler.setOptions(curr_month,self.entry_monthList,self.entry_month)
        
    def check_database_thread(self) -> None:
        """ Database thread to check database (Might be overkill) """
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.fetch_data,self.quit)
        table: Optional[pd.DataFrame] = None
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                table = self.QUEUE.get()
                self.clear_queue()
                break
            
        if(table is not None):
            GUI_Handler.view_excel(table,self.outer.CHILD[DataView.__name__].text_excel) # type: ignore
            unique_col = mapping(table.columns,'HR EMP CODE')
            
            jsons = PDF_TEMPLATE.check_json()
            htmls = PDF_TEMPLATE.check_html()
            
            if(htmls and jsons):
                GUI_Handler.setOptions(jsons,self.outer.CHILD[DataView.__name__].json_list,self.outer.CHILD[DataView.__name__].json) # type: ignore
                GUI_Handler.setOptions(htmls,self.outer.CHILD[DataView.__name__].html_list,self.outer.CHILD[DataView.__name__].html) # type: ignore
                
                if(unique_col is not None):
                    BaseTemplate.data = table
                    self.outer.CHILD[DataView.__name__].id_column = unique_col # type: ignore
                    self.switch_screen(DataView)
                    tkmb.showinfo('Fetch Status','Data Fetched Successfully')
                else:
                    tkmb.showinfo('Fetch Status','Data does not have HR EMP Code Column in table')
            else:
                tkmb.showerror('Fetch Status','Template HTML/Mapping JSON files were not found!')
                
        else:
            tkmb.showinfo('Fetch Status','Data Does not exists')
            
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
            
    def cancel_thread_wrapper(self) -> None:
        self.cancel_thread()
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        tkmb.showinfo('Fetch Status','Data Fetching Stopped')
        
    # fetches data from the chosen table
    def getData(self) -> None:
        curr_institute = self.entry_institute.get()
        curr_type = self.entry_type.get()
        curr_year = self.entry_year.get()
        curr_month = self.entry_month.get()
        
        self.outer.CHILD[DataView.__name__].chosen_month = curr_month # type: ignore
        self.outer.CHILD[DataView.__name__].chosen_year = curr_year # type: ignore
        self.outer.CHILD[DataView.__name__].chosen_type = curr_type # type: ignore
        self.outer.CHILD[DataView.__name__].chosen_institute = curr_institute # type: ignore
        self.outer.CHILD[DataView.__name__].changeHeading() # type: ignore
        
        if (self.can_start_thread()):
            self.process = Process(target=DatabaseWrapper(**self.outer.CRED).get_data,kwargs={'queue':self.QUEUE,'institute': curr_institute,'type': curr_type,'year': curr_year,'month': curr_month},daemon=True)
            self.thread = Thread(target=self.check_database_thread,daemon=True)
            self.process.start()
            self.thread.start()
            
        else:
            tkmb.showerror("Processing Error","Cannot fetch data from Database now. Please try again")
            
class DataView(BaseTemplate):
    
    chosen_month = "None"
    chosen_year = "None"
    chosen_type = "None"
    chosen_institute = "None"
    size:int = 12
    id_column: NullStr = None
    html = ctk.StringVar(value="index.html")
    json = ctk.StringVar(value="index.json")
    
    def __init__(self,outer:App):

        super().__init__(outer)
        self.outer = outer
        master = self.outer.APP
        
        self.frame = ctk.CTkScrollableFrame(master=master, fg_color=COLOR_SCHEME["fg_color"],)
        
        self.label_date = ctk.CTkLabel(master=self.frame , text=f"Data View for {self.chosen_institute.capitalize()} {self.chosen_type.capitalize()} {self.chosen_month.capitalize()}/{self.chosen_year.upper()}", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250)
        self.label_date.pack(pady=20,padx=10)
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Enter Employee ID:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_id = ctk.CTkEntry(master=frame , text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.entry_id.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        self.clipboard = ctk.CTkButton(master=self.frame , text="Copy Row to Clipboard", command=self.copy_row_to_clipboard, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.clipboard.pack(pady=10)

        ctk.CTkLabel(master=self.frame , text="Choose Template HTML and Mapping JSON: ", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=50).pack(padx=10)
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Template HTML:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.html_list = ctk.CTkOptionMenu(master=frame,variable=self.html,values=[],button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.html_list.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Mapping JSON:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.json_list = ctk.CTkOptionMenu(master=frame,variable=self.json,values=[],button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.json_list.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        self.single_print = ctk.CTkButton(master=self.frame , text="Generate Single PDF", command=self.single_print_pdf_cover, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.single_print.pack(pady=10)
        
        self.bulk_print = ctk.CTkButton(master=self.frame , text="Bulk Print PDFs", command=self.bulk_print_pdfs_cover, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.bulk_print.pack(pady=10)
        
        self.back = ctk.CTkButton(master=self.frame , text='Back', command=self.back_to_preview, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back.pack(pady=10, padx=10)
        
        ctk.CTkLabel(master=self.frame , text="Text Size of Sheet:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10,padx=10)
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        
        self.minus = ctk.CTkButton(master=frame , text="-", command=self.decrease_size, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50,state="disabled")
        self.minus.pack(side="left",padx=10)
        
        self.row = ctk.CTkLabel(master=frame , text=f"{self.size}", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=100)
        self.row.pack(side="left",padx=10)
        
        self.plus = ctk.CTkButton(master=frame , text="+", command=self.increase_size, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.plus.pack(side="left",padx=10)  
        
        frame.pack()

        self.text_excel = scrolledtext.ScrolledText(master=self.frame, width=90, height=25,  bg="black", fg="white", wrap=tk.NONE, font=("Courier", 12))
        x_scrollbar = Scrollbar(self.frame, orient="horizontal",command=self.text_excel.xview)
        self.text_excel.pack(pady=10,padx=10, fill='both', expand=True)
        x_scrollbar.pack(side='bottom', fill='x')

        self.text_excel.configure(xscrollcommand=x_scrollbar.set)
        self.quit = ctk.CTkButton(master=self.frame, text='Quit the Process',command=self.stop_pdf_thread, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.to_disable = list(self.get_widgets_to_disable())

    # back to interface
    def changeHeading(self):
        GUI_Handler.changeText(self.label_date,f"Data View for {self.chosen_institute.capitalize()} {self.chosen_type.capitalize()} {self.chosen_month.capitalize()}/{self.chosen_year}")
    
    def back_to_preview(self) -> None:
        self.clear_data()
        self.clear_queue()
        GUI_Handler.clear_excel(self.text_excel)
        self.hide()
        self.outer.CHILD[DataPreview.__name__].appear()
        
    def decrease_size(self) -> None:
        self.row.configure(text=max(MIN_TEXT_SIZE,self.size-1))
        self.size = max(MIN_TEXT_SIZE,self.size-1)
        
        GUI_Handler.change_text_font(self.text_excel,self.size)
        GUI_Handler.unlock_gui_button([self.plus])
        if(self.size==MIN_TEXT_SIZE): GUI_Handler.lock_gui_button([self.minus])
        
    def increase_size(self) -> None:
        self.row.configure(text=min(MAX_TEXT_SIZE,self.size+1))
        self.size = min(MAX_TEXT_SIZE,self.size+1)
        
        GUI_Handler.change_text_font(self.text_excel,self.size)
        GUI_Handler.unlock_gui_button([self.minus])
        if(self.size==MAX_TEXT_SIZE): GUI_Handler.lock_gui_button([self.plus])
    
    def copy_to_clipboard_thread(self, emp_id:str) -> None:
        
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.clipboard,self.quit)
        
        if(BaseTemplate.data is None):
            tkmb.showwarning("Data Error", "Data was not found")
            return
        
        search_result = BaseTemplate.data[BaseTemplate.data[str(self.id_column)]==emp_id]
        
        if(not search_result.empty): 
            pyperclip.copy(','.join(map(text_clean,search_result.iloc[[0]].to_numpy())))
            tkmb.showinfo("ClipBoard Status", "Employee data copied to clipboard.")
        else:
            tkmb.showwarning("Error", "Employee ID was not found.")

        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
    def copy_row_to_clipboard(self) -> None:
        emp_id = text_clean(self.entry_id.get())
        
        if(emp_id):
            if(self.can_start_thread()):
                self.thread = Thread(target=self.copy_to_clipboard_thread,kwargs={'emp_id':emp_id},daemon=True)
                self.thread.start()
                
            else:
                tkmb.showerror('Program Status',f"Warning Background Thread is still running")
        else:
            tkmb.showerror("Empty Data","Empty Search String Detected")

    def single_pdf_thread(self) -> None:
        
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.single_print,self.quit)
        status:bool | NullStr = None
        msg: NullStr = None
        
        while True:
            
            if(self.stop_flag):return None
            
            if(not self.QUEUE.empty()):
                curr = self.QUEUE.get()
                
                if(isinstance(curr,tuple) and len(curr)==2):
                    status, msg = curr
                    
                self.clear_queue()
                break
            
        if(status and msg):
            tkmb.showinfo('Single PDF Status',msg)
        elif(msg):
            tkmb.showwarning('Single PDF Status',msg)
        else:
            tkmb.showwarning('Single PDF Status',"Something went wrong. Please try again")
            
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
    def stop_pdf_thread(self):
        self.cancel_thread()
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
        tkmb.showinfo('Process Status','PDF Generation Cancelled')

    # cover for extract_data
    def single_print_pdf_cover(self) -> None:
        month = self.chosen_month
        year = self.chosen_year
        emp_id = text_clean(self.entry_id.get())
        json = self.json.get()
        html = self.html.get()
        
        if(not emp_id):
            tkmb.showerror("Empty Data","Empty Search String Detected")
            return
        
        if(BaseTemplate.data is None):
            tkmb.showerror("Data Error","Data is Unavailable")
            return
        
        file: NullStr = None
        
        try:
            _file = filedialog.asksaveasfile(initialfile=f"employee_{emp_id}",defaultextension=".pdf",initialdir=Path(APP_PATH).parent,filetypes=[("PDF File","*.pdf")])
            
            if _file is not None:
                file = _file.name
                _file.close()
                
        except Exception as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
            tkmb.showwarning("Error",f"Some error has occured: {e}")
            file = None
    
        if file:
            os.remove(file)
            if(self.can_start_thread()):
                PDF_TEMPLATE.chosen_json = json
                PDF_TEMPLATE.chosen_html = html
                self.process = Process(target=PandaWrapper(BaseTemplate.data,str(self.id_column),PDF_TEMPLATE).litany_of_scroll,kwargs={'queue':self.QUEUE,'data_slate_path': Path(file),'month': month,'year': year,'emp_id': emp_id},daemon=True)
                self.thread = Thread(target=self.single_pdf_thread,daemon=True)
                self.process.start()
                self.thread.start()
                
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
                
        else:
            tkmb.showerror("File Error","File was not found")
            
                        
    def bulk_print_pdfs_thread(self) -> None:
        
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.bulk_print,self.quit)
        done:int = 0
        total:int = 0
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                hot = self.QUEUE.get()
                
                if(isinstance(hot, tuple) and len(hot)==2):
                    done, total = hot
                    
                self.clear_queue()
                break
                
        if((not self.stop_flag) and total):
            tkmb.showinfo('Bulk PDF Status',f"Generated {done} PDFs out of {total} records")
        elif(not self.stop_flag):
            tkmb.showwarning('Bulk PDF Status',f"No PDFs were generated")
            
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
    def bulk_print_pdfs_cover(self) -> None:
                
        month = self.chosen_month
        year = self.chosen_year
        institute = self.chosen_institute
        type = self.chosen_type
        json = self.json.get()
        html = self.html.get()
        where: Optional[Path] = None
        
        try:
            where =  Path(APP_PATH).parent
            
            where = where.joinpath('pdfs', institute, type, year, month)
            os.makedirs(where.resolve())

        except OSError as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))

        if where is None:
            tkmb.showerror('Program Status',f"File path was invalid")

        if(self.can_start_thread()):
            PDF_TEMPLATE.chosen_html = Path(html)
            PDF_TEMPLATE.chosen_json = Path(json)
            self.process = Process(target=PandaWrapper(BaseTemplate.data,str(self.id_column),PDF_TEMPLATE).litany_of_scrolls,kwargs={'queue': self.QUEUE,'dropzone': where,'month': month,'year': year},daemon=True)
            self.thread = Thread(target=self.bulk_print_pdfs_thread,daemon=True)
            
            self.process.start()
            self.thread.start()
            
        else:
            tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")

class FileInput(BaseTemplate):
    
    sheet = ctk.StringVar(value='Sheet1')
    row_index:int = 0
    size:int = 12
    max_row:int = 0
    encryption:bool = False
    prev_password: str = ''
    
    def __init__(self, outer:App):
        super().__init__(outer)
        self.outer = outer
        master = self.outer.APP
        
        self.frame = ctk.CTkScrollableFrame(master=master, fg_color=COLOR_SCHEME["fg_color"],)

        # title
        ctk.CTkLabel(master=self.frame , text=f"File Upload", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)

        # file path input
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Select Excel File:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.file = ctk.CTkEntry(master=frame , text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.file.pack(padx=10,pady=10,side='left')
        frame.pack()

        # browse button
        self.browse_button = ctk.CTkButton(master=self.frame , text="Browse", command=self.select_file, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.browse_button.pack(pady=10)
        
        self.upload_button = ctk.CTkButton(master=self.frame , text="Upload", command=self.load_decrypted_file, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back = ctk.CTkButton(master=self.frame , text='Back', command=self.back_to_interface, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back.pack(pady=10, padx=10)
        
        # password frame
        self.password_frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=self.password_frame, text="Enter Password for File:",text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.password_box = ctk.CTkEntry(master=self.password_frame ,show='*',text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.password_box.pack(padx=10,pady=10,side='left')

        self.variable_frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        """ frames are added after file is uploaded """
        
        
        frame = ctk.CTkFrame(master=self.variable_frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Sheet Present:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.sheetList = ctk.CTkOptionMenu(master=frame,variable=self.sheet,values=[],command=self.changeView,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.sheetList.pack(side='left',pady=10, padx=10)
        frame.pack()
        
        ctk.CTkLabel(master=self.variable_frame , text="Row Change:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10,padx=10)
        
        frame = ctk.CTkFrame(master=self.variable_frame, fg_color=COLOR_SCHEME["fg_color"])
        self.row_minus = ctk.CTkButton(master=frame , text="-", command=self.prev_row, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.row_minus.pack(side="left",padx=10)
        self.row = ctk.CTkLabel(master=frame , text=f"{self.row_index}", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=150)
        self.row.pack(side="left",padx=10)
        self.row_plus = ctk.CTkButton(master=frame , text="+", command=self.next_row, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.row_plus.pack(side="left",padx=10)        
        frame.pack(pady=10, padx=10)
        
        self.upload = ctk.CTkButton(master=self.variable_frame , text="Save to DB", command=self.go_to_upload, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.upload.pack(pady=10)
        self.variable_back = ctk.CTkButton(master=self.variable_frame , text='Back', command=self.back_to_interface, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.variable_back.pack(pady=10, padx=10)

        ctk.CTkLabel(master=self.variable_frame , text="Text Size of Sheet:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10,padx=10)
        
        frame = ctk.CTkFrame(master=self.variable_frame, fg_color=COLOR_SCHEME["fg_color"])
        
        self.size_minus = ctk.CTkButton(master=frame , text="-", command=self.decrease_size, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50,state="disabled")
        self.size_minus.pack(side="left",padx=10)
        self.font_size = ctk.CTkLabel(master=frame , text=f"{self.size}", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=100)
        self.font_size.pack(side="left",padx=10)
        self.size_plus = ctk.CTkButton(master=frame , text="+", command=self.increase_size, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.size_plus.pack(side="left",padx=10)  
        frame.pack()

        self.text_excel = scrolledtext.ScrolledText(master=self.variable_frame, width=300, height=50,  bg="black", fg="white", wrap=tk.NONE, font=("Courier", 12))
        x_scrollbar = Scrollbar(self.variable_frame, orient="horizontal",command=self.text_excel.xview)
        x_scrollbar.pack(side='bottom', fill='x')

        self.text_excel.configure(xscrollcommand=x_scrollbar.set)
        self.text_excel.pack(pady=10,padx=10, fill='both', expand=True)
        
        self.quit = ctk.CTkButton(master=self.frame, text='Quit the Process',command=self.cancel_thread_wrapper, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.to_disable = list(self.get_widgets_to_disable())
        
        
    def back_to_interface(self):
        GUI_Handler.clear_excel(self.text_excel)
        self.set_for_file_upload_state()
        self.clear_data()
        self.switch_screen(Interface)
        
    def go_to_upload(self):
        """ passing data """
        if(self.can_start_thread()):
            self.thread = Thread(target=self._go_to_upload_thread,daemon=True)
            self.thread.start()
        else:
            tkmb.showerror('Program Status',f"Warning Background Thread is still running")
        
    def _go_to_upload_thread(self):
        GUI_Handler.view_excel(BaseTemplate.data[self.sheet.get()],self.outer.CHILD[UploadData.__name__].text_excel)        
        self.outer.CHILD[UploadData.__name__].sheet = self.sheet.get()
        self.switch_screen(UploadData)
        
    def cancel_thread_wrapper(self):
        self.cancel_thread()
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
    def next_row(self):
        max_row = BaseTemplate.data[self.sheet.get()].shape[0]
        
        self.row_index = min(max_row, self.row_index+1)
        
        GUI_Handler.changeText(self.row,self.row_index)
        if(self.row_index==max_row): GUI_Handler.lock_gui_button([self.row_plus])
        else:
            GUI_Handler.unlock_gui_button([self.row_minus])
        self.get_data()

    def prev_row(self):
        
        self.row_index = max(0, self.row_index-1)
        GUI_Handler.changeText(self.row,self.row_index)
        
        if(self.row_index==0): GUI_Handler.lock_gui_button([self.row_minus])
        else:
            GUI_Handler.unlock_gui_button([self.row_plus])
        self.get_data()

    def get_data(self):
        file_path = self.file.get()
        password = self.prev_password
        
        if(file_path):
            
            if(self.encryption and password):                
                if(self.can_start_thread()):
                    self.thread = Thread(target=self.change_row_thread,daemon=True)
                    self.process = Process(target=Decryption.fetch_encrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'password':password,'skip':self.row_index},daemon=True)
                    self.thread.start()
                    self.process.start()
                    
                else:
                    tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
                    
            elif(self.encryption):
                tkmb.showinfo('Encryption Status',f"File '{file_path}' is encrypted. Please enter the password")
                
            elif(not self.encryption):
                if(self.can_start_thread()):
                    self.thread = Thread(target=self.change_row_thread,daemon=True)
                    self.process = Process(target=Decryption.fetch_decrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'skip':self.row_index},daemon=True)
                    self.process.start()
                    self.thread.start()
                    
                else:
                    tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
        else:
            tkmb.showerror('File Status',f"Empty file path was detected")

    def decrease_size(self):
        self.font_size.configure(text=max(MIN_TEXT_SIZE,self.size-1))
        self.size = max(MIN_TEXT_SIZE,self.size-1)
        
        GUI_Handler.change_text_font(self.text_excel,self.size)
        GUI_Handler.unlock_gui_button([self.size_plus])
        if(self.size==MIN_TEXT_SIZE): GUI_Handler.lock_gui_button([self.size_minus])
        
    def increase_size(self) -> None:
        self.font_size.configure(text=min(MAX_TEXT_SIZE,self.size+1))
        self.size = min(MAX_TEXT_SIZE,self.size+1)
        
        GUI_Handler.change_text_font(self.text_excel,self.size)
        GUI_Handler.unlock_gui_button([self.size_minus])
        if(self.size==MAX_TEXT_SIZE): GUI_Handler.lock_gui_button([self.size_plus])
        
    def encryption_check_thread(self):
        """ Locks present GUI and places quit button after browse button """
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.browse_button,self.quit)  
        
    def set_after_upload_state(self):
        """ Places variable frame and removes upload and password frame """
        GUI_Handler.remove_widget(self.upload_button)
        GUI_Handler.remove_widget(self.password_frame)
        GUI_Handler.place_after(self.browse_button,self.variable_frame)
        GUI_Handler.remove_widget(self.back)
        
    def set_after_file_is_encrypted_state(self):
        """ Places password frame before upload button """
        GUI_Handler.place_before(self.upload_button,self.password_frame)
        
    def set_for_file_upload_state(self):
        """ Removes password and variable frame """
        GUI_Handler.remove_widget(self.variable_frame)
        GUI_Handler.remove_widget(self.password_frame)
        GUI_Handler.remove_widget(self.upload_button)
        GUI_Handler.place_after(self.browse_button, self.back)
        GUI_Handler.clear_entry(self.file)
        GUI_Handler.clear_entry(self.password_box)

    def is_encrypted_thread(self):
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.browse_button,self.quit)
        self.set_for_file_upload_state()  
        GUI_Handler.changeCommand(self.upload_button,self.load_decrypted_file)
        
        is_encrypted: bool | None = None
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                is_encrypted = self.QUEUE.get()
                self.clear_queue()
                break
            
        self.encryption = is_encrypted
        
        if(is_encrypted is not None):
            GUI_Handler.place_after(self.browse_button,self.upload_button)
            
            if(is_encrypted):
                self.set_after_file_is_encrypted_state()
                GUI_Handler.changeCommand(self.upload_button,self.load_encrypted_file)
                
                tkmb.showinfo('Encrypted Status',f"Excel File '{self.file.get()}' is encrypted. Please provide the password")
        else:    
            tkmb.showerror('File Status',f"File '{self.file.get()}' was not found")
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
        
    def select_file(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", ".xlsx;.xls")],initialdir=Path(APP_PATH).parent)

        if(file_path):
            GUI_Handler.change_file_holder(self.file,file_path)
            
            if(self.can_start_thread()):
                self.thread = Thread(target=self.is_encrypted_thread,daemon=True)
                self.process = Process(target=Decryption.is_encrypted_wrapper,kwargs={'queue': self.QUEUE,'file_path': Path(file_path)},daemon=True)
                self.thread.start()
                self.process.start()
                
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
                
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")
            
    
    def load_unprotected_data_thread(self) -> None:
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.upload_button,self.quit)
        self.set_for_file_upload_state()  
        
        sheets: list[str] = []
        data: Optional[pd.DataFrame] = None
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                sheets, data = self.QUEUE.get()
                self.clear_queue()
                break
            
        
        if((sheets) and (data is not None)):
            BaseTemplate.data = data
            GUI_Handler.setOptions(sheets,self.sheetList,self.sheet)
            self.set_after_upload_state()
            GUI_Handler.view_excel(BaseTemplate.data[sheets[0]],self.text_excel)
            
            tkmb.showinfo('Upload Status',f"Excel File '{self.file.get()}' was loaded")
        else:
            
            tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' could not be loaded")
            
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
    
    
    def load_decrypted_file(self):
        file_path = self.file.get()
                
        if(file_path):
            
            if(self.can_start_thread()):
                self.thread = Thread(target=self.load_unprotected_data_thread,daemon=True)
                self.process = Process(target=Decryption.fetch_decrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'skip':self.row_index},daemon=True)
                self.thread.start()
                self.process.start()
                
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
        
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")
    
    
    def load_protected_data_thread(self) -> None:
        
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.upload_button,self.quit)
        
        sheets:list[str] = []
        data: Optional[pd.DataFrame] = None
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                sheets, data = self.QUEUE.get()
                self.clear_queue()
                break

        if((sheets) and (data is not None)):
            BaseTemplate.data = data
            GUI_Handler.setOptions(sheets,self.sheetList,self.sheet)
            self.set_after_upload_state()
            GUI_Handler.view_excel(data[sheets[0]],self.text_excel)
            
            self.prev_password = self.password_box.get()
            tkmb.showinfo('Upload Status',f"Excel File '{self.file.get()}' was loaded")
        else:
            
            tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' could not be loaded. Please check the password")
            
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
    def load_encrypted_file(self):
        file_path = self.file.get()
        password = self.password_box.get()
        
        if(file_path):
            
            if(self.encryption and (not password)):
                tkmb.showinfo('Encryption Status',f"File '{file_path}' is encrypted. Please enter the password")
                return None
                
            if(self.can_start_thread()):
                self.thread = Thread(target=self.load_protected_data_thread,daemon=True)
                self.process = Process(target=Decryption.fetch_encrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'password':password,'skip':self.row_index},daemon=True)
                self.thread.start()
                self.process.start()
            
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
        
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")
            
    
    def change_view_thread(self):
        GUI_Handler.lock_gui_button(self.to_disable)
        current_sheet = self.sheet.get()
        current_data = BaseTemplate.data[current_sheet]
        
        GUI_Handler.view_excel(current_data,self.text_excel)
        GUI_Handler.unlock_gui_button(self.to_disable)
        
    def change_row_thread(self):
        
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_before(self.text_excel,self.quit)
        current_sheet = self.sheet.get()
        data: Optional[dict[str, pd.DataFrame]] = None
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                _, data = self.QUEUE.get()
                self.clear_queue()
                break
        
        if((current_sheet is not None) and (data is not None)):
            BaseTemplate.data = data
            self.set_after_upload_state()
            GUI_Handler.view_excel(data[current_sheet],self.text_excel)
            self.prev_password = self.password_box.get()
        else:
            
            tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' could not be loaded. Please check the password")
        
        GUI_Handler.view_excel(BaseTemplate.data[self.sheet.get()],self.text_excel)
        GUI_Handler.unlock_gui_button(self.to_disable)    
        GUI_Handler.remove_widget(self.quit)
        
    
    def changeView(self, event: Optional[tk.Event] = None):
        sheet = self.sheet.get()
        
        if(sheet):
            
            if(self.can_start_thread()):
                self.thread = Thread(target=self.change_view_thread,daemon=True)
                self.thread.start()
            else:
                tkmb.showerror('Program Status',f"Warning Background Thread is still running")
        
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")

class UploadData(BaseTemplate):
    size:int = 12        
    chosen_institute = ctk.StringVar(value='Somaiya')
    chosen_type = ctk.StringVar(value='Teaching')
    chosen_month = ctk.StringVar(value='Jan')  
    sheet: NullStr = None

    def __init__(self, outer:App):
        super().__init__(outer)

        # title
        ctk.CTkLabel(master=self.frame , text=f"Data Upload", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)

        # file path input
        ctk.CTkLabel(master=self.frame, text="Please Enter Details about data", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 18, "bold")).pack(padx=10,pady=10)

        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Institute:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_institute =ctk.CTkOptionMenu(master=frame,variable=self.chosen_institute,values=list(self.outer.TOGGLE.keys()),command=self.changeType,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.entry_institute.pack(pady=10, padx=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Type:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.toggle_type = ctk.CTkOptionMenu(master=frame,variable=self.chosen_type,values=self.outer.TOGGLE[list(self.outer.TOGGLE.keys())[0]],button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.toggle_type.pack(pady=10, padx=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Month:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_month = ctk.CTkOptionMenu(master=frame,variable=self.chosen_month,values=[str(i).capitalize() for i in self.outer.MONTH],button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.entry_month.pack(pady=10, padx=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame , text="Enter Year:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(side='left',pady=10, padx=10)
        self.entry_year = ctk.CTkEntry(master=frame ,placeholder_text="Eg. 2024", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=100)
        self.entry_year.pack(side='left',pady=10, padx=10)
        frame.pack()
        
        self.create_button = ctk.CTkButton(master=self.frame , text='Create Table', command=self.create_in_db, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.create_button.pack(pady=10, padx=10)
        
        self.update_button = ctk.CTkButton(master=self.frame , text='Upload To DB', command=self.update_in_db, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        
        self.delete_button = ctk.CTkButton(master=self.frame , text='Delete from DB', command=self.delete_from_db, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.delete_button.pack(pady=10, padx=10)
        
        self.back = ctk.CTkButton(master=self.frame , text='Back', command=self.back_to_input, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back.pack(pady=10, padx=10)
        
        ctk.CTkLabel(master=self.frame , text="Text Size of Sheet:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10,padx=10)
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        
        self.size_minus = ctk.CTkButton(master=frame , text="-", command=self.decrease_size, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.size_minus.pack(side="left",padx=10)
        self.font_size = ctk.CTkLabel(master=frame , text=f"{self.size}", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=100)
        self.font_size.pack(side="left",padx=10)
        self.size_plus = ctk.CTkButton(master=frame , text="+", command=self.increase_size, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.size_plus.pack(side="left",padx=10)  
        frame.pack()
        
        self.text_excel = scrolledtext.ScrolledText(master=self.frame, width=300, height=50,  bg="black", fg="white", wrap=tk.NONE, font=("Courier", 12))
        x_scrollbar = Scrollbar(self.frame, orient="horizontal",command=self.text_excel.xview)
        x_scrollbar.pack(side='bottom', fill='x')

        self.text_excel.configure(xscrollcommand=x_scrollbar.set)
        self.text_excel.pack(pady=10,padx=10, fill='both', expand=True)
        
        
        self.quit = ctk.CTkButton(master=self.frame, text='Quit the Process',command=self.cancel_thread_wrapper, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.to_disable = list(self.get_widgets_to_disable())
        

    def back_to_input(self):
        self.clear_data(hard=False)
        GUI_Handler.remove_widget(self.update_button)
        GUI_Handler.clear_excel(self.text_excel)
        self.switch_screen(FileInput)
        
    def cancel_thread_wrapper(self):
        self.cancel_thread()
        GUI_Handler.unlock_gui_button(self.to_disable)    
        GUI_Handler.remove_widget(self.quit)
        
    def decrease_size(self):
        self.font_size.configure(text=max(MIN_TEXT_SIZE,self.size-1))
        self.size = max(MIN_TEXT_SIZE,self.size-1)
        
        GUI_Handler.change_text_font(self.text_excel,self.size)
        GUI_Handler.unlock_gui_button([self.size_plus])
        if(self.size==MIN_TEXT_SIZE): GUI_Handler.lock_gui_button([self.size_minus])
        
    def increase_size(self):
        self.font_size.configure(text=max(MAX_TEXT_SIZE,self.size+1))
        self.size = max(MAX_TEXT_SIZE,self.size+1)
        
        GUI_Handler.change_text_font(self.text_excel,self.size)
        GUI_Handler.unlock_gui_button([self.size_minus])
        if(self.size==MAX_TEXT_SIZE): GUI_Handler.lock_gui_button([self.size_plus])

    def changeType(self, event: Optional[tk.Event] = None):
        institute = self.chosen_institute.get()
        type = self.outer.TOGGLE[institute]
        GUI_Handler.setOptions(type,self.toggle_type,self.chosen_type)
        
    def create_thread(self,month:str,year:int,institute:str,type:str):
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.create_button,self.quit)
        result:int = 0
    
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                result = self.QUEUE.get()
                self.clear_queue()
                break
                
        if(result is not None):
            match result:
                case CreateTable.NO_ID: tkmb.showwarning('Column Info', CreateTable.NO_ID)
                case CreateTable.SUCCESS: tkmb.showinfo("Database Status", CreateTable.SUCCESS)
                case CreateTable.COLUMNS_MISMATCH: tkmb.showinfo("Database Status", CreateTable.COLUMNS_MISMATCH)
                case CreateTable.ERROR: tkmb.showerror("Database Status", CreateTable.ERROR)
                case CreateTable.EXISTS: tkmb.showerror("Database Status", CreateTable.EXISTS)
                
            if(result in {CreateTable.EXISTS, CreateTable.SUCCESS}):
                GUI_Handler.place_after(self.create_button,self.update_button)
                
        else:
            tkmb.showinfo("Database Status", f"MySQL Error occurred")
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
                
    
    def create_in_db(self):
        institute = self.chosen_institute.get()
        type = self.chosen_type.get()
        month = self.chosen_month.get()
        year = self.entry_year.get()

        if(self.sheet is None):
            tkmb.showwarning('Excel Sheet', 'Excel Sheet was not chosen')
        
        if(year_check(year)):
            if(self.can_start_thread() and messagebox.askyesnocancel("Confirmation", f"Month: {month}, Year: {year} \n Institue: {institute}, Type: {type} \n Are you sure details are correct?")):
                self.thread = Thread(target=self.create_thread,kwargs={'month': month,'year': year,'institute': institute,'type': type},daemon=True)
                self.process = Process(target=DatabaseWrapper(**self.outer.CRED).create_table,kwargs={'queue': self.QUEUE,'institute': institute,'type': type,'year': year,'month': month,'data_columns': list(BaseTemplate.data[self.sheet].columns)},daemon=True)
                self.thread.start()
                self.process.start()
                
        else:
            tkmb.showwarning("Alert","Incorrect year format")

    def update_thread(self):
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.update_button,self.quit)
        result:int = 0
    
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                result = self.QUEUE.get()
                self.clear_queue()
                break
            
        if(result is not None):
            match result:
                case UpdateTable.COLUMNS_MISMATCH: tkmb.showwarning('Column Info', UpdateTable.COLUMNS_MISMATCH)
                case UpdateTable.NO_ID: tkmb.showwarning('Column Info', UpdateTable.NO_ID)
                case UpdateTable.ERROR: tkmb.showinfo("Database Status", UpdateTable.ERROR)
                case UpdateTable.SUCCESS: tkmb.showinfo("Database Status", UpdateTable.SUCCESS)
                
        else:
            tkmb.showinfo("Database Status", f"MySQL Error occurred")
                
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)

    def update_in_db(self):
        institute = self.chosen_institute.get()
        type = self.chosen_type.get()
        month = self.chosen_month.get()
        year = self.entry_year.get()
        data = BaseTemplate.data[self.sheet]

        if(year_check(year)):
            if(self.can_start_thread() and messagebox.askyesnocancel("Upload Confirmation", "Would you like it upload this data. Any pre-existing data will be updated. Are you sure about this?")):
                self.thread = Thread(target=self.update_thread,daemon=True)
                self.process = Process(target=DatabaseWrapper(**self.outer.CRED).fill_table,kwargs={'data':data,'queue': self.QUEUE,'institute': institute,'type': type,'year': year,'month': month},daemon=True)
                self.thread.start()
                self.process.start()
                
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
                
        else:
            tkmb.showwarning("Alert","Incorrect year format")

    def delete_thread(self) -> None:
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.delete_button,self.quit)
        result:bool | None = None
    
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                result = self.QUEUE.get()
                self.clear_queue()
                break
            
        
        match(result):
            case DeleteTable.ERROR: tkmb.showwarning("Database Status", DeleteTable.ERROR)
            case DeleteTable.TABLE_NOT_FOUND: tkmb.showwarning("Database Status", DeleteTable.TABLE_NOT_FOUND)
            case DeleteTable.SUCCESS: tkmb.showinfo("Database Status", DeleteTable.SUCCESS)
                
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        

    def delete_from_db(self):
        institute = self.chosen_institute.get()
        type = self.chosen_type.get()
        month = self.chosen_month.get()
        year = self.entry_year.get()
        
        
        if(year_check(year)):
            if(self.can_start_thread() and messagebox.askyesnocancel("Confirmation", f"Month: {month}, Year: {year} \n Institute:{institute}, Type: {type} \n Are you sure you want to clear this data from DB?")):
                self.thread = Thread(target=self.delete_thread,daemon=True)
                self.process = Process(target=DatabaseWrapper(**self.outer.CRED).delete_table,kwargs={'queue': self.QUEUE,'institute': institute,'type': type,'year': year,'month': month},daemon=True)
                self.thread.start()
                self.process.start()
                
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
                
        else:
            tkmb.showwarning("Alert","Incorrect year format")

class DataPeek(BaseTemplate):
    """ Previews data in database """
    entry_year = ctk.StringVar(value='2024')
    entry_month = ctk.StringVar(value='Jan')
    entry_institute = ctk.StringVar(value='Somaiya')
    entry_type = ctk.StringVar(value='Teaching')
    tables:dict[str,dict[str,dict[str,set[str]]]] = {}
    
    def __init__(self,outer:App) -> None:
        super().__init__(outer)
        self.outer = outer
        master = self.outer.APP
        self.frame = ctk.CTkScrollableFrame(master=master, fg_color=COLOR_SCHEME["fg_color"],)

        ctk.CTkLabel(master=self.frame , text="Data Preview", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Institute:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_instituteList = ctk.CTkOptionMenu(master=frame,variable=self.entry_institute,values=[],command=self.changeInstitute,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.entry_instituteList.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Type:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_typeList = ctk.CTkOptionMenu(master=frame,variable=self.entry_type,values=[],command=self.changeType,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.entry_typeList.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Year:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_yearList = ctk.CTkOptionMenu(master=frame,variable=self.entry_year,values=[],command=self.changeYear,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.entry_yearList.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Month:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.entry_monthList = ctk.CTkOptionMenu(master=frame,variable=self.entry_month,values=[],button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.entry_monthList.pack(padx=10,pady=10,side='left')
        frame.pack()
        
        self.fetch_data = ctk.CTkButton(master=self.frame , text='Continue', command=self.go_to_delete, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.fetch_data.pack(pady=10, padx=10)
        
        self.back = ctk.CTkButton(master=self.frame , text='Back', command=self.back_to_interface, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back.pack(pady=10, padx=10)
        
        self.quit = ctk.CTkButton(master=self.frame,text='Quit the process',command=self.cancel_thread_wrapper,fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.to_disable = list(self.get_widgets_to_disable())
        
    def back_to_interface(self) -> None:
        """ Moves back to the earlier page """
        self.hide()
        self.tables = {}
        self.outer.CHILD[Interface.__name__].appear()
            
    def changeData(self):
        """ Initialization of the optionmenus """     
        curr_institute = list(self.tables.keys())
        curr_type = list(self.tables[curr_institute[0]].keys())
        curr_year = list(self.tables[curr_institute[0]][curr_type[0]].keys())
        curr_month = sorted(list(self.tables[curr_institute[0]][curr_type[0]][curr_year[0]]),key=lambda x:self.outer.MONTH[x])
        
        GUI_Handler.setOptions(curr_type,self.entry_typeList,self.entry_type)
        GUI_Handler.setOptions(curr_institute,self.entry_instituteList,self.entry_institute)
        GUI_Handler.setOptions(curr_year,self.entry_yearList,self.entry_year)
        GUI_Handler.setOptions(curr_month,self.entry_monthList,self.entry_month)
        
    def changeInstitute(self,event:tk.Event) -> None:
        """ Changes Institute Type """
        curr_institute = self.entry_institute.get()
        curr_type = list(self.tables[curr_institute].keys())
        curr_year = list(self.tables[curr_institute][curr_type[0]].keys())
        curr_month = sorted(list(self.tables[curr_institute][curr_type[0]][curr_year[0]]),key=lambda x:self.outer.MONTH[x])

        GUI_Handler.setOptions(curr_type,self.entry_typeList,self.entry_type)
        GUI_Handler.setOptions(curr_year,self.entry_yearList,self.entry_year)
        GUI_Handler.setOptions(curr_month,self.entry_monthList,self.entry_month)
        
    def changeType(self,event:tk.Event):
        """ Changes Type Type """
        curr_institute = self.entry_institute.get()
        curr_type = self.entry_type.get()
        curr_year = list(self.tables[curr_institute][curr_type].keys())
        curr_month = sorted(list(self.tables[curr_institute][curr_type][curr_year[0]]),key=lambda x:self.outer.MONTH[x])
        
        GUI_Handler.setOptions(curr_year,self.entry_yearList,self.entry_year)
        GUI_Handler.setOptions(curr_month,self.entry_monthList,self.entry_month)
        
    # updates menu when year changes
    def changeYear(self,event:tk.Event) -> None:
        """ Changes the year """
        curr_institute = self.entry_institute.get()
        curr_type = self.entry_type.get()
        curr_year = self.entry_year.get()
        curr_month = sorted(list(self.tables[curr_institute][curr_type][curr_year]),key=lambda x:self.outer.MONTH[x])

        GUI_Handler.setOptions(curr_month,self.entry_monthList,self.entry_month)
        
    def cancel_thread_wrapper(self) -> None:
        self.cancel_thread()
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        tkmb.showinfo('Fetch Status','Data Fetching Stopped')
        
    # fetches data from the chosen table
    def go_to_delete(self) -> None:
        curr_institute = self.entry_institute.get()
        curr_type = self.entry_type.get()
        curr_year = self.entry_year.get()
        curr_month = self.entry_month.get()
        
        self.outer.CHILD[DeleteView.__name__].chosen_month = curr_month # type: ignore
        self.outer.CHILD[DeleteView.__name__].chosen_year = curr_year # type: ignore
        self.outer.CHILD[DeleteView.__name__].chosen_type = curr_type # type: ignore
        self.outer.CHILD[DeleteView.__name__].chosen_institute = curr_institute # type: ignore
        self.outer.CHILD[DeleteView.__name__].changeHeading() # type: ignore
        
        self.hide()
        self.outer.CHILD[DeleteView.__name__].appear()
            
    def preview(self) -> None:
        
        if(self.can_start_thread()):
            self.process = Process(target=DatabaseWrapper(**self.outer.CRED).check_table,kwargs={'queue':self.QUEUE},daemon=True)
            self.thread = Thread(target=self.check_database_thread,daemon=True)

            self.thread.start()
            self.process.start()
            
        else:
            tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
            
    def check_database_thread(self):
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.fetch_data,self.quit)
        table = {}
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                table = self.QUEUE.get()
                self.clear_queue()
                break
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
        if(table):
            self.tables = table
            self.changeData()
        else:
            tkmb.showerror("MySQL Error","No Data Available to delete")
            self.hide()
            self.outer.CHILD[Interface.__name__].appear()
    
class DeleteView(BaseTemplate):
    
    chosen_month = "None"
    chosen_year = "None"
    chosen_type = "None"
    chosen_institute = "None"
    
    def __init__(self,outer:App):

        super().__init__(outer)
        self.outer = outer
        master = self.outer.APP
        
        self.frame = ctk.CTkScrollableFrame(master=master, fg_color=COLOR_SCHEME["fg_color"],)
        
        self.label_date = ctk.CTkLabel(master=self.frame , text=f"Data View for {self.chosen_institute.capitalize()} {self.chosen_type.capitalize()} {self.chosen_month.capitalize()}/{self.chosen_year.upper()}", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250)
        self.label_date.pack(pady=20,padx=10)
                
        self.delete_button = ctk.CTkButton(master=self.frame , text="Delete Table", command=self.delete_from_db, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.delete_button.pack(pady=10)
                
        self.back = ctk.CTkButton(master=self.frame , text='Back', command=self.back_to_view, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back.pack(pady=10, padx=10)
        
        self.quit = ctk.CTkButton(master=self.frame, text='Quit the Process',command=self.stop_delete_thread, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.to_disable = list(self.get_widgets_to_disable())
        

    # back to interface
    def changeHeading(self):
        GUI_Handler.changeText(self.label_date,f"Data View for {self.chosen_institute.capitalize()} {self.chosen_type.capitalize()} {self.chosen_month.capitalize()}/{self.chosen_year}")
    
    def back_to_view(self) -> None:
        self.clear_data(True)
        self.clear_queue()
        self.switch_screen(DataPeek)
        self.outer.CHILD[DataPeek.__name__].preview() # type: ignore
        
    def stop_delete_thread(self):
        self.cancel_thread()
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
        tkmb.showinfo('Process Status','PDF Generation Cancelled')
    
    def delete_thread(self) -> None:
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.delete_button,self.quit)
        result:bool | None = None
    
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                result = self.QUEUE.get()
                self.clear_queue()
                break
            
        
        if(result):
            tkmb.showinfo("Database Status", f"Table dropped successfully")
        elif(result is not None):
            tkmb.showinfo("Database Status", f"Table does not exists")
        else:
            tkmb.showinfo("Database Status", f"MySQL Error occured")
                
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
    def delete_from_db(self):
        institute = self.chosen_institute
        type = self.chosen_type
        month = self.chosen_month
        year = self.chosen_year
        
        if(year_check(year)):
            if(self.can_start_thread() and messagebox.askyesnocancel("Confirmation", f"Month: {month}, Year: {year} \n Institue:{institute}, Type: {type} \n Are you sure you want to clear this data from DB?")):
                self.thread = Thread(target=self.delete_thread,daemon=True)
                self.process = Process(target=DatabaseWrapper(**self.outer.CRED).delete_table,kwargs={'queue': self.QUEUE,'institute': institute,'type': type,'year': year,'month': month},daemon=True)
                self.thread.start()
                self.process.start()
                
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
                
        else:
            tkmb.showwarning("Alert","Incorrect year format")

class TemplateInput(BaseTemplate):
    
    sheet = ctk.StringVar(value='Sheet1')
    row_index:int = 0
    size:int = 12
    max_row:int = 0
    encryption: bool = False
    prev_password: str = ''
    
    def __init__(self, outer:App):
        super().__init__(outer)

        # title
        ctk.CTkLabel(master=self.frame , text=f"Template Generator", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)

        # file path input
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Select Excel File:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.file = ctk.CTkEntry(master=frame , text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.file.pack(padx=10,pady=10,side='left')
        frame.pack()

        # browse button
        self.browse_button = ctk.CTkButton(master=self.frame , text="Browse", command=self.select_file, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.browse_button.pack(pady=10)
        
        self.upload_button = ctk.CTkButton(master=self.frame , text="Upload", command=self.load_decrypted_file, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back = ctk.CTkButton(master=self.frame , text='Back', command=self.back_to_interface, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back.pack(pady=10, padx=10)
        
        # password frame
        self.password_frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=self.password_frame, text="Enter Password for File:",text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.password_box = ctk.CTkEntry(master=self.password_frame ,show='*',text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.password_box.pack(padx=10,pady=10,side='left')

        self.variable_frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        """ frames are added after file is uploaded """
        
        frame = ctk.CTkFrame(master=self.variable_frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Sheet Present:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.sheetList = ctk.CTkOptionMenu(master=frame,variable=self.sheet,values=[],command=self.changeView,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.sheetList.pack(side='left',pady=10, padx=10)
        frame.pack(pady=10, padx=10)
        
        ctk.CTkLabel(master=self.variable_frame , text="Row Change:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10,padx=10)
        
        frame = ctk.CTkFrame(master=self.variable_frame, fg_color=COLOR_SCHEME["fg_color"])
        self.row_minus = ctk.CTkButton(master=frame , text="-", command=self.prev_row, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.row_minus.pack(side="left",padx=10)
        self.row = ctk.CTkLabel(master=frame , text=f"{self.row_index}", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=150)
        self.row.pack(side="left",padx=10)
        self.row_plus = ctk.CTkButton(master=frame , text="+", command=self.next_row, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.row_plus.pack(side="left",padx=10)        
        frame.pack(pady=10, padx=10)
        
        frame = ctk.CTkFrame(master=self.variable_frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame , text="Enter Template Name:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(side='left',pady=10, padx=10)
        self.template = ctk.CTkEntry(master=frame ,placeholder_text="Eg. Somaiya_Template", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=100)
        self.template.pack(side='left',pady=10, padx=10)
        frame.pack(pady=10, padx=10)
        
        self.upload = ctk.CTkButton(master=self.variable_frame , text="Generate Template", command=self.generate_template, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.upload.pack(pady=10)
        self.variable_back = ctk.CTkButton(master=self.variable_frame , text='Back', command=self.back_to_interface, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.variable_back.pack(pady=10, padx=10)

        ctk.CTkLabel(master=self.variable_frame , text="Text Size of Sheet:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10,padx=10)
        
        frame = ctk.CTkFrame(master=self.variable_frame, fg_color=COLOR_SCHEME["fg_color"])
        
        self.size_minus = ctk.CTkButton(master=frame , text="-", command=self.decrease_size, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50,state="disabled")
        self.size_minus.pack(side="left",padx=10)
        self.font_size = ctk.CTkLabel(master=frame , text=f"{self.size}", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=100)
        self.font_size.pack(side="left",padx=10)
        self.size_plus = ctk.CTkButton(master=frame , text="+", command=self.increase_size, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.size_plus.pack(side="left",padx=10)  
        frame.pack()

        self.text_excel = scrolledtext.ScrolledText(master=self.variable_frame, width=300, height=50,  bg="black", fg="white", wrap=tk.NONE, font=("Courier", 12))
        x_scrollbar = Scrollbar(self.variable_frame, orient="horizontal",command=self.text_excel.xview)
        x_scrollbar.pack(side='bottom', fill='x')

        self.text_excel.configure(xscrollcommand=x_scrollbar.set)
        self.text_excel.pack(pady=10,padx=10, fill='both', expand=True)
        
        self.quit = ctk.CTkButton(master=self.frame, text='Quit the Process',command=self.cancel_thread_wrapper, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.to_disable = list(self.get_widgets_to_disable())
        
        
    def back_to_interface(self):
        GUI_Handler.clear_excel(self.text_excel)
        self.set_for_file_upload_state()
        self.clear_data()
        self.switch_screen(Interface)
        
    def generate_template(self):
        """ passing data """
        
        file_name = self.template.get()
        
        if(not file_name):
            tkmb.showerror("Template Generation", "File Name cannot be empty!")
            return
        
        if(BaseTemplate.data is None):
            tkmb.showerror("Template Generation","Data is Unavailable")
            return
        
        if(self.can_start_thread()):
            self.process = Process(target=TemplateGenerator.make_template, kwargs={"file_name": file_clean(file_name), "data": BaseTemplate.data, "queue": self.QUEUE}, daemon=True)
            self.thread = Thread(target=self.go_to_generate_template_thread,daemon=True)
            self.process.start()
            self.thread.start()
        else:
            tkmb.showerror('Program Status',f"Warning Background Thread is still running")
        
    def go_to_generate_template_thread(self) -> None:
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.upload, self.quit)
        
        status: Optional[bool] = None
        msg: str = "Something went wrong"
        counterThat: int = 2
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                result = self.QUEUE.get()
                
                if(isinstance(result, tuple) and len(result)==2):
                    status, msg = result
                    
                    if(status is not None):
                        if(status):
                            tkmb.showinfo("Template Generation", msg)
                        else:
                            tkmb.showwarning("Template Generation", msg)
                            
                    else:
                        tkmb.showinfo("Template Generation", f"Template Generation Failed")
                
                counterThat -= 1
                
            if counterThat == 0:
                self.clear_queue()
                break
                
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
    def cancel_thread_wrapper(self):
        self.cancel_thread()
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
    def next_row(self):
        max_row = BaseTemplate.data[self.sheet.get()].shape[0]
        
        self.row_index = min(max_row, self.row_index+1)
        
        GUI_Handler.changeText(self.row,self.row_index)
        if(self.row_index==max_row): GUI_Handler.lock_gui_button([self.row_plus])
        else:
            self.get_data()
            GUI_Handler.unlock_gui_button([self.row_minus])

    def prev_row(self):
        
        self.row_index = max(0, self.row_index-1)
        GUI_Handler.changeText(self.row,self.row_index)
        
        if(self.row_index==0): GUI_Handler.lock_gui_button([self.row_minus])
        else:
            self.get_data()
            GUI_Handler.unlock_gui_button([self.row_plus])

    def get_data(self):
        file_path = self.file.get()
        password = self.prev_password
        
        if(file_path):
            
            if(self.encryption and password):                
                if(self.can_start_thread()):
                    self.thread = Thread(target=self.change_row_thread,daemon=True)
                    self.process = Process(target=Decryption.fetch_encrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'password':password,'skip':self.row_index},daemon=True)
                    self.thread.start()
                    self.process.start()
                    
                else:
                    tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
                    
            elif(self.encryption):
                tkmb.showinfo('Encryption Status',f"File '{file_path}' is encrypted. Please enter the password")
                
            elif(not self.encryption):
                if(self.can_start_thread()):
                    self.thread = Thread(target=self.change_row_thread,daemon=True)
                    self.process = Process(target=Decryption.fetch_decrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'skip':self.row_index},daemon=True)
                    self.process.start()
                    self.thread.start()
                    
                else:
                    tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
        else:
            tkmb.showerror('File Status',f"Empty file path was detected")

    def decrease_size(self):
        self.font_size.configure(text=max(MIN_TEXT_SIZE,self.size-1))
        self.size = max(MIN_TEXT_SIZE,self.size-1)
        
        GUI_Handler.change_text_font(self.text_excel,self.size)
        GUI_Handler.unlock_gui_button([self.size_plus])
        if(self.size==MIN_TEXT_SIZE): GUI_Handler.lock_gui_button([self.size_minus])
        
    def increase_size(self) -> None:
        self.font_size.configure(text=min(MAX_TEXT_SIZE,self.size+1))
        self.size = min(MAX_TEXT_SIZE,self.size+1)
        
        GUI_Handler.change_text_font(self.text_excel,self.size)
        GUI_Handler.unlock_gui_button([self.size_minus])
        if(self.size==MAX_TEXT_SIZE): GUI_Handler.lock_gui_button([self.size_plus])
        
    def encryption_check_thread(self):
        """ Locks present GUI and places quit button after browse button """
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.browse_button,self.quit)  
        
    def set_after_upload_state(self):
        """ Places variable frame and removes upload and password frame """
        GUI_Handler.remove_widget(self.upload_button)
        GUI_Handler.remove_widget(self.password_frame)
        GUI_Handler.place_after(self.browse_button,self.variable_frame)
        GUI_Handler.remove_widget(self.back)
        
    def set_after_file_is_encrypted_state(self):
        """ Places password frame before upload button """
        GUI_Handler.place_before(self.upload_button,self.password_frame)
        
    def set_for_file_upload_state(self):
        """ Removes password and variable frame """
        GUI_Handler.remove_widget(self.variable_frame)
        GUI_Handler.remove_widget(self.password_frame)
        GUI_Handler.remove_widget(self.upload_button)
        GUI_Handler.place_after(self.browse_button, self.back)
        GUI_Handler.clear_entry(self.file)
        GUI_Handler.clear_entry(self.password_box)

    def is_encrypted_thread(self) -> None:
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.browse_button,self.quit)
        self.set_for_file_upload_state()  
        GUI_Handler.changeCommand(self.upload_button,self.load_decrypted_file)
        
        is_encrypted: bool | None = None
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                is_encrypted = self.QUEUE.get()
                self.clear_queue()
                break
            
        self.encryption = is_encrypted
        
        if(is_encrypted is not None):
            GUI_Handler.place_after(self.browse_button,self.upload_button)
            
            if(is_encrypted):
                self.set_after_file_is_encrypted_state()
                GUI_Handler.changeCommand(self.upload_button,self.load_encrypted_file)
                
                tkmb.showinfo('Encrypted Status',f"Excel File '{self.file.get()}' is encrypted. Please provide the password")
        else:    
            tkmb.showerror('File Status',f"File '{self.file.get()}' was not found")
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
        
    def select_file(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", ".xlsx;.xls")],initialdir=Path(APP_PATH).parent)

        if(file_path):
            GUI_Handler.change_file_holder(self.file,file_path)
            
            if(self.can_start_thread()):
                self.thread = Thread(target=self.is_encrypted_thread,daemon=True)
                self.process = Process(target=Decryption.is_encrypted_wrapper,kwargs={'queue': self.QUEUE,'file_path': Path(file_path)},daemon=True)
                self.thread.start()
                self.process.start()
                
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
                
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")
            
    
    def load_unprotected_data_thread(self) -> None:
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.upload_button,self.quit)
        self.set_for_file_upload_state()  
        
        try:
            sheets: list[str] = []
            data: Optional[dict[str, pd.DataFrame]] = None
            
            while True:
                
                if(self.stop_flag): return None
                
                if(not self.QUEUE.empty()):
                    sheets, data = self.QUEUE.get()
                    self.clear_queue()
                    break
                
            
            if((sheets) and (data is not None)):
                
                if(not checkColumns(sheets, TEMPLATE_SHEET)):
                    tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' requires all the following sheets: {', '.join(TEMPLATE_SHEET)}")
                    return
                
                if(not all([checkColumns(_data.columns, TEMPLATE_COLUMN) for _data in data.values()])):
                    tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' requires all the following columns in all sheets: {', '.join(TEMPLATE_COLUMN)}")
                    return
                
                BaseTemplate.data = data
                GUI_Handler.setOptions(sheets,self.sheetList,self.sheet)
                self.set_after_upload_state()
                GUI_Handler.view_excel(BaseTemplate.data[sheets[0]],self.text_excel)
                
                tkmb.showinfo('Upload Status',f"Excel File '{self.file.get()}' was loaded")
            else:
                
                tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' could not be loaded. Please ensure file")
        finally:        
            GUI_Handler.unlock_gui_button(self.to_disable)
            GUI_Handler.remove_widget(self.quit)
    
    def load_decrypted_file(self):
        file_path = self.file.get()
                
        if(file_path):
            
            if(self.can_start_thread()):
                self.thread = Thread(target=self.load_unprotected_data_thread,daemon=True)
                self.process = Process(target=Decryption.fetch_decrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'skip':self.row_index},daemon=True)
                self.thread.start()
                self.process.start()
                
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
        
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")
    
    
    def load_protected_data_thread(self) -> None:
        
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.upload_button,self.quit)
        
        try:
            sheets:list[str] = []
            data: Optional[dict[str, pd.DataFrame]] = None
            
            while True:
                
                if(self.stop_flag): return None
                
                if(not self.QUEUE.empty()):
                    sheets, data = self.QUEUE.get()
                    self.clear_queue()
                    break

            if((sheets) and (data is not None)):

                if(not checkColumns(sheets, TEMPLATE_SHEET)):
                    tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' requires all the following sheets: {', '.join(TEMPLATE_SHEET)}")
                    return
                
                if(not all([checkColumns(_data.columns, TEMPLATE_COLUMN) for _data in data.values()])):
                    tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' requires all the following columns in all sheets: {', '.join(TEMPLATE_COLUMN)}")
                    return
                    
                BaseTemplate.data = data
                GUI_Handler.setOptions(sheets,self.sheetList,self.sheet)
                self.set_after_upload_state()
                GUI_Handler.view_excel(data[sheets[0]],self.text_excel)
                
                self.prev_password = self.password_box.get()
                tkmb.showinfo('Upload Status',f"Excel File '{self.file.get()}' was loaded")
            else:
                
                tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' could not be loaded. Please check the password")
                
        finally:
            GUI_Handler.unlock_gui_button(self.to_disable)
            GUI_Handler.remove_widget(self.quit)
        
    def load_encrypted_file(self):
        file_path = self.file.get()
        password = self.password_box.get()
        
        if(file_path):
            
            if(self.encryption and (not password)):
                tkmb.showinfo('Encryption Status',f"File '{file_path}' is encrypted. Please enter the password")
                return None
                
            if(self.can_start_thread()):
                self.thread = Thread(target=self.load_protected_data_thread,daemon=True)
                self.process = Process(target=Decryption.fetch_encrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'password':password,'skip':self.row_index},daemon=True)
                self.thread.start()
                self.process.start()
            
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
        
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")
            
    
    def change_view_thread(self):
        GUI_Handler.lock_gui_button(self.to_disable)
        current_sheet = self.sheet.get()
        current_data = BaseTemplate.data[current_sheet]
        
        GUI_Handler.view_excel(current_data,self.text_excel)
        GUI_Handler.unlock_gui_button(self.to_disable)
        
    def change_row_thread(self) -> None:
        
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_before(self.text_excel,self.quit)
        current_sheet = self.sheet.get()
        data: Optional[dict[str, pd.DataFrame]] = None
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                _, data = self.QUEUE.get()
                self.clear_queue()
                break
        
        if((current_sheet is not None) and (data is not None)):
            BaseTemplate.data = data
            self.set_after_upload_state()
            GUI_Handler.view_excel(data[current_sheet],self.text_excel)
            self.prev_password = self.password_box.get()
        else:
            
            tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' could not be loaded. Please check the password")
        
        GUI_Handler.view_excel(BaseTemplate.data[self.sheet.get()],self.text_excel)
        GUI_Handler.unlock_gui_button(self.to_disable)    
        GUI_Handler.remove_widget(self.quit)
        
    
    def changeView(self, event: Optional[tk.Event] = None):
        sheet = self.sheet.get()
        
        if(sheet):
            
            if(self.can_start_thread()):
                self.thread = Thread(target=self.change_view_thread,daemon=True)
                self.thread.start()
            else:
                tkmb.showerror('Program Status',f"Warning Background Thread is still running")
        
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")

class TemplateGeneration(BaseTemplate):
    
    sheet = ctk.StringVar(value='Sheet1')
    row_index:int = 0
    size:int = 12
    max_row:int = 0
    encryption: bool = False
    prev_password: str = ''
    
    def __init__(self, outer:App):
        super().__init__(outer)

        # title
        ctk.CTkLabel(master=self.frame , text=f"Template Download", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 25, "bold"),width=250).pack(pady=20,padx=10)

        # file path input
        frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Select Excel File:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.file = ctk.CTkEntry(master=frame , text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.file.pack(padx=10,pady=10,side='left')
        frame.pack()

        # browse button
        self.browse_button = ctk.CTkButton(master=self.frame , text="Browse", command=self.select_file, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.browse_button.pack(pady=10)
        
        self.upload_button = ctk.CTkButton(master=self.frame , text="Upload", command=self.load_decrypted_file, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back = ctk.CTkButton(master=self.frame , text='Back', command=self.back_to_interface, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.back.pack(pady=10, padx=10)
        
        # password frame
        self.password_frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=self.password_frame, text="Enter Password for File:",text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.password_box = ctk.CTkEntry(master=self.password_frame ,show='*',text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.password_box.pack(padx=10,pady=10,side='left')

        self.variable_frame = ctk.CTkFrame(master=self.frame, fg_color=COLOR_SCHEME["fg_color"])
        """ frames are added after file is uploaded """
        
        frame = ctk.CTkFrame(master=self.variable_frame, fg_color=COLOR_SCHEME["fg_color"])
        ctk.CTkLabel(master=frame, text="Sheet Present:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold")).pack(padx=10,pady=10,side='left')
        self.sheetList = ctk.CTkOptionMenu(master=frame,variable=self.sheet,values=[],command=self.changeView,button_color=COLOR_SCHEME["button_color"],fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=200)
        self.sheetList.pack(side='left',pady=10, padx=10)
        frame.pack(pady=10, padx=10)
        
        ctk.CTkLabel(master=self.variable_frame , text="Row Change:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10,padx=10)
        
        frame = ctk.CTkFrame(master=self.variable_frame, fg_color=COLOR_SCHEME["fg_color"])
        self.row_minus = ctk.CTkButton(master=frame , text="-", command=self.prev_row, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.row_minus.pack(side="left",padx=10)
        self.row = ctk.CTkLabel(master=frame , text=f"{self.row_index}", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=150)
        self.row.pack(side="left",padx=10)
        self.row_plus = ctk.CTkButton(master=frame , text="+", command=self.next_row, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.row_plus.pack(side="left",padx=10)        
        frame.pack(pady=10, padx=10)
        
        self.upload = ctk.CTkButton(master=self.variable_frame , text="Download Template", command=self.generate_template, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.upload.pack(pady=10)
        self.variable_back = ctk.CTkButton(master=self.variable_frame , text='Back', command=self.back_to_interface, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.variable_back.pack(pady=10, padx=10)

        ctk.CTkLabel(master=self.variable_frame , text="Text Size of Sheet:", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=250).pack(pady=10,padx=10)
        
        frame = ctk.CTkFrame(master=self.variable_frame, fg_color=COLOR_SCHEME["fg_color"])
        
        self.size_minus = ctk.CTkButton(master=frame , text="-", command=self.decrease_size, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50,state="disabled")
        self.size_minus.pack(side="left",padx=10)
        self.font_size = ctk.CTkLabel(master=frame , text=f"{self.size}", text_color=COLOR_SCHEME["text_color"], font=("Ubuntu", 16, "bold"),width=100)
        self.font_size.pack(side="left",padx=10)
        self.size_plus = ctk.CTkButton(master=frame , text="+", command=self.increase_size, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=50)
        self.size_plus.pack(side="left",padx=10)  
        frame.pack()

        self.text_excel = scrolledtext.ScrolledText(master=self.variable_frame, width=300, height=50,  bg="black", fg="white", wrap=tk.NONE, font=("Courier", 12))
        x_scrollbar = Scrollbar(self.variable_frame, orient="horizontal",command=self.text_excel.xview)
        x_scrollbar.pack(side='bottom', fill='x')

        self.text_excel.configure(xscrollcommand=x_scrollbar.set)
        self.text_excel.pack(pady=10,padx=10, fill='both', expand=True)
        
        self.quit = ctk.CTkButton(master=self.frame, text='Quit the Process',command=self.cancel_thread_wrapper, fg_color=COLOR_SCHEME["button_color"], font=("Ubuntu", 16, "bold"),width=250)
        self.to_disable = list(self.get_widgets_to_disable())
        
        
    def back_to_interface(self):
        GUI_Handler.clear_excel(self.text_excel)
        self.set_for_file_upload_state()
        self.clear_data()
        self.switch_screen(Interface)
        
    def generate_template(self) -> None:
        """ passing data """
        
        file_name = self.file.get()
        
        if(BaseTemplate.data is None):
            tkmb.showerror("Template Generation","Data is Unavailable")
            return
        
        file: NullStr = None
        
        try:
            _file = filedialog.asksaveasfile(initialfile=f"template.xlsx",defaultextension=".xlsx",initialdir=Path(APP_PATH).parent,filetypes=[("XLSX File","*.xlsx")])
            
            if _file is not None:
                file = _file.name
                _file.close()
                
        except Exception as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
            file = None
        
        if(not ((file_name))):
            tkmb.showerror("Template Generation", "File Name cannot be empty!")
            return
        
        if (file is None):
            tkmb.showerror("Template Generation", "Destination File cannot be empty!")
            return
        
        if(self.can_start_thread()):
            self.process = Process(target=TemplateGenerator.make_excel, kwargs={"file_name": Path(file), "data": BaseTemplate.data, "queue": self.QUEUE}, daemon=True)
            self.thread = Thread(target=self.go_to_generate_template_thread,daemon=True)
            self.process.start()
            self.thread.start()
        else:
            tkmb.showerror('Program Status',f"Warning Background Thread is still running")
        
    def go_to_generate_template_thread(self) -> None:
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.upload, self.quit)
        
        status: Optional[bool] = None
        msg: str = "Something went wrong"
        counterThat: int = 1
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                result = self.QUEUE.get()
                
                if(isinstance(result, tuple) and len(result)==2):
                    status, msg = result
                    
                    if(status is not None):
                        if(status):
                            tkmb.showinfo("Template Generation", msg)
                        else:
                            tkmb.showwarning("Template Generation", msg)
                            
                    else:
                        tkmb.showinfo("Template Generation", f"Template Generation Failed")
                
                counterThat -= 1
                
            if counterThat == 0:
                self.clear_queue()
                break
                
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
    def cancel_thread_wrapper(self):
        self.cancel_thread()
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
    def next_row(self):
        max_row = BaseTemplate.data[self.sheet.get()].shape[0]
        
        self.row_index = min(max_row, self.row_index+1)
        
        GUI_Handler.changeText(self.row,self.row_index)
        if(self.row_index==max_row): GUI_Handler.lock_gui_button([self.row_plus])
        else:
            self.get_data()
            GUI_Handler.unlock_gui_button([self.row_minus])

    def prev_row(self):
        
        self.row_index = max(0, self.row_index-1)
        GUI_Handler.changeText(self.row,self.row_index)
        
        if(self.row_index==0): GUI_Handler.lock_gui_button([self.row_minus])
        else:
            self.get_data()
            GUI_Handler.unlock_gui_button([self.row_plus])

    def get_data(self):
        file_path = self.file.get()
        password = self.prev_password
        
        if(file_path):
            
            if(self.encryption and password):                
                if(self.can_start_thread()):
                    self.thread = Thread(target=self.change_row_thread,daemon=True)
                    self.process = Process(target=Decryption.fetch_encrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'password':password,'skip':self.row_index},daemon=True)
                    self.thread.start()
                    self.process.start()
                    
                else:
                    tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
                    
            elif(self.encryption):
                tkmb.showinfo('Encryption Status',f"File '{file_path}' is encrypted. Please enter the password")
                
            elif(not self.encryption):
                if(self.can_start_thread()):
                    self.thread = Thread(target=self.change_row_thread,daemon=True)
                    self.process = Process(target=Decryption.fetch_decrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'skip':self.row_index},daemon=True)
                    self.process.start()
                    self.thread.start()
                    
                else:
                    tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
        else:
            tkmb.showerror('File Status',f"Empty file path was detected")

    def decrease_size(self):
        self.font_size.configure(text=max(MIN_TEXT_SIZE,self.size-1))
        self.size = max(MIN_TEXT_SIZE,self.size-1)
        
        GUI_Handler.change_text_font(self.text_excel,self.size)
        GUI_Handler.unlock_gui_button([self.size_plus])
        if(self.size==MIN_TEXT_SIZE): GUI_Handler.lock_gui_button([self.size_minus])
        
    def increase_size(self) -> None:
        self.font_size.configure(text=min(MAX_TEXT_SIZE,self.size+1))
        self.size = min(MAX_TEXT_SIZE,self.size+1)
        
        GUI_Handler.change_text_font(self.text_excel,self.size)
        GUI_Handler.unlock_gui_button([self.size_minus])
        if(self.size==MAX_TEXT_SIZE): GUI_Handler.lock_gui_button([self.size_plus])
        
    def encryption_check_thread(self):
        """ Locks present GUI and places quit button after browse button """
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.browse_button,self.quit)  
        
    def set_after_upload_state(self):
        """ Places variable frame and removes upload and password frame """
        GUI_Handler.remove_widget(self.upload_button)
        GUI_Handler.remove_widget(self.password_frame)
        GUI_Handler.place_after(self.browse_button,self.variable_frame)
        GUI_Handler.remove_widget(self.back)
        
    def set_after_file_is_encrypted_state(self):
        """ Places password frame before upload button """
        GUI_Handler.place_before(self.upload_button,self.password_frame)
        
    def set_for_file_upload_state(self):
        """ Removes password and variable frame """
        GUI_Handler.remove_widget(self.variable_frame)
        GUI_Handler.remove_widget(self.password_frame)
        GUI_Handler.remove_widget(self.upload_button)
        GUI_Handler.place_after(self.browse_button, self.back)
        GUI_Handler.clear_entry(self.file)
        GUI_Handler.clear_entry(self.password_box)

    def is_encrypted_thread(self) -> None:
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.browse_button,self.quit)
        self.set_for_file_upload_state()  
        GUI_Handler.changeCommand(self.upload_button,self.load_decrypted_file)
        
        is_encrypted: bool | None = None
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                is_encrypted = self.QUEUE.get()
                self.clear_queue()
                break
            
        self.encryption = is_encrypted
        
        if(is_encrypted is not None):
            GUI_Handler.place_after(self.browse_button,self.upload_button)
            
            if(is_encrypted):
                self.set_after_file_is_encrypted_state()
                GUI_Handler.changeCommand(self.upload_button,self.load_encrypted_file)
                
                tkmb.showinfo('Encrypted Status',f"Excel File '{self.file.get()}' is encrypted. Please provide the password")
        else:    
            tkmb.showerror('File Status',f"File '{self.file.get()}' was not found")
        
        GUI_Handler.unlock_gui_button(self.to_disable)
        GUI_Handler.remove_widget(self.quit)
        
        
    def select_file(self) -> None:
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", ".xlsx;.xls")],initialdir=Path(APP_PATH).parent)

        if(file_path):
            GUI_Handler.change_file_holder(self.file,file_path)
            
            if(self.can_start_thread()):
                self.thread = Thread(target=self.is_encrypted_thread,daemon=True)
                self.process = Process(target=Decryption.is_encrypted_wrapper,kwargs={'queue': self.QUEUE,'file_path': Path(file_path)},daemon=True)
                self.thread.start()
                self.process.start()
                
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
                
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")
            
    
    def load_unprotected_data_thread(self) -> None:
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.upload_button,self.quit)
        self.set_for_file_upload_state()  
        
        try:
            sheets: list[str] = []
            data: Optional[dict[str, pd.DataFrame]] = None
            
            while True:
                
                if(self.stop_flag): return None
                
                if(not self.QUEUE.empty()):
                    sheets, data = self.QUEUE.get()
                    self.clear_queue()
                    break
                
            
            if((sheets) and (data is not None)):
                
                if(not checkColumns(sheets, TEMPLATE_SHEET)):
                    tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' requires all the following sheets: {', '.join(TEMPLATE_SHEET)}")
                    return
                
                if(not all([checkColumns(_data.columns, TEMPLATE_COLUMN) for _data in data.values()])):
                    tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' requires all the following columns in all sheets: {', '.join(TEMPLATE_COLUMN)}")
                    return
                
                BaseTemplate.data = data
                GUI_Handler.setOptions(sheets,self.sheetList,self.sheet)
                self.set_after_upload_state()
                GUI_Handler.view_excel(BaseTemplate.data[sheets[0]],self.text_excel)
                
                tkmb.showinfo('Upload Status',f"Excel File '{self.file.get()}' was loaded")
            else:
                
                tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' could not be loaded. Please ensure file")
        finally:        
            GUI_Handler.unlock_gui_button(self.to_disable)
            GUI_Handler.remove_widget(self.quit)
    
    def load_decrypted_file(self):
        file_path = self.file.get()
                
        if(file_path):
            
            if(self.can_start_thread()):
                self.thread = Thread(target=self.load_unprotected_data_thread,daemon=True)
                self.process = Process(target=Decryption.fetch_decrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'skip':self.row_index},daemon=True)
                self.thread.start()
                self.process.start()
                
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
        
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")
    
    
    def load_protected_data_thread(self) -> None:
        
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_after(self.upload_button,self.quit)
        
        try:
            sheets:list[str] = []
            data: Optional[dict[str, pd.DataFrame]] = None
            
            while True:
                
                if(self.stop_flag): return None
                
                if(not self.QUEUE.empty()):
                    sheets, data = self.QUEUE.get()
                    self.clear_queue()
                    break

            if((sheets) and (data is not None)):

                if(not checkColumns(sheets, TEMPLATE_SHEET)):
                    tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' requires all the following sheets: {', '.join(TEMPLATE_SHEET)}")
                    return
                
                if(not all([checkColumns(_data.columns, TEMPLATE_COLUMN) for _data in data.values()])):
                    tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' requires all the following columns in all sheets: {', '.join(TEMPLATE_COLUMN)}")
                    return
                    
                BaseTemplate.data = data
                GUI_Handler.setOptions(sheets,self.sheetList,self.sheet)
                self.set_after_upload_state()
                GUI_Handler.view_excel(data[sheets[0]],self.text_excel)
                
                self.prev_password = self.password_box.get()
                tkmb.showinfo('Upload Status',f"Excel File '{self.file.get()}' was loaded")
            else:
                
                tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' could not be loaded. Please check the password")
                
        finally:
            GUI_Handler.unlock_gui_button(self.to_disable)
            GUI_Handler.remove_widget(self.quit)
        
    def load_encrypted_file(self):
        file_path = self.file.get()
        password = self.password_box.get()
        
        if(file_path):
            
            if(self.encryption and (not password)):
                tkmb.showinfo('Encryption Status',f"File '{file_path}' is encrypted. Please enter the password")
                return None
                
            if(self.can_start_thread()):
                self.thread = Thread(target=self.load_protected_data_thread,daemon=True)
                self.process = Process(target=Decryption.fetch_encrypted_file,kwargs={'queue': self.QUEUE,'file_path': Path(file_path),'password':password,'skip':self.row_index},daemon=True)
                self.thread.start()
                self.process.start()
            
            else:
                tkmb.showerror('Program Status',f"Warning Background Process/Thread is still running")
        
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")
            
    
    def change_view_thread(self):
        GUI_Handler.lock_gui_button(self.to_disable)
        current_sheet = self.sheet.get()
        current_data = BaseTemplate.data[current_sheet]
        
        GUI_Handler.view_excel(current_data,self.text_excel)
        GUI_Handler.unlock_gui_button(self.to_disable)
        
    def change_row_thread(self) -> None:
        
        GUI_Handler.lock_gui_button(self.to_disable)
        GUI_Handler.place_before(self.text_excel,self.quit)
        current_sheet = self.sheet.get()
        data: Optional[dict[str, pd.DataFrame]] = None
        
        while True:
            
            if(self.stop_flag): return None
            
            if(not self.QUEUE.empty()):
                _, data = self.QUEUE.get()
                self.clear_queue()
                break
        
        if((current_sheet is not None) and (data is not None)):
            BaseTemplate.data = data
            self.set_after_upload_state()
            GUI_Handler.view_excel(data[current_sheet],self.text_excel)
            self.prev_password = self.password_box.get()
        else:
            
            tkmb.showwarning('File Status',f"Excel File '{self.file.get()}' could not be loaded. Please check the password")
        
        GUI_Handler.view_excel(BaseTemplate.data[self.sheet.get()],self.text_excel)
        GUI_Handler.unlock_gui_button(self.to_disable)    
        GUI_Handler.remove_widget(self.quit)
        
    
    def changeView(self, event: Optional[tk.Event] = None):
        sheet = self.sheet.get()
        
        if(sheet):
            
            if(self.can_start_thread()):
                self.thread = Thread(target=self.change_view_thread,daemon=True)
                self.thread.start()
            else:
                tkmb.showerror('Program Status',f"Warning Background Thread is still running")
        
        else:
            tkmb.showerror('File Status',f"Empty file_path was detected")


class TkErrorCatcher:
    """ Logging Tkinter Loops errors """

    def __init__(self, func, subst, widget):
        self.func = func
        self.subst = subst
        self.widget = widget

    def __call__(self, *args):
        try:
            if self.subst:
                args = self.subst(*args)
            return self.func(*args)
        
        except SystemExit as msg:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(msg))
            raise SystemExit(msg)
        
        except Exception as e:
            ERROR_LOG.write_error(ERROR_LOG.get_error_info(e))
            raise Exception(e)
        
def show_error(exctype:type[BaseException], excvalue:BaseException, tb:types.TracebackType, thread: NullStr = None):
    if (issubclass(exctype, KeyboardInterrupt)):
        sys.__excepthook__(exctype, excvalue, tb)
        return

    err = "{} \n {}".format(ERROR_LOG.get_error_info(exctype(excvalue)),'\n'.join(traceback.format_exception(exctype, excvalue, tb))) # type: ignore
    
    if(thread is None):
        ERROR_LOG.write_error(err)
    else:
        ERROR_LOG.write_error(err,'APPLICATION:THREAD')

sys.excepthook = show_error # type: ignore
excepthook = show_error # type: ignore

if __name__ == "__main__":
    # initialize main app
    freeze_support()
    tk.CallWrapper = TkErrorCatcher # type: ignore
    app = App(500,500,'Excel-To-Pdf Generator')
    app.APP.report_callback_exception = show_error
    PDF_TEMPLATE.load_default()
    app.start_app()
    app.exit_app()