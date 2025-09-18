import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
from logger import Logger
import aiosmtplib
from type import *
from pathlib import Path

class Mailing():
    """ a class for Mailing (supports method-chaining) """

    def __init__(self, email:str, key:str, error_log:Logger) -> None:
        self.email = email
        self.key = key
        self.error = error_log
        self.msg = MIMEMultipart()
        self.msg['From'] = email
        self.smtp: Optional[smtplib.SMTP] = None
        self.status = True
    
    def addTxtMsg(self, msg:str, msgType:str) -> 'Mailing':
        """ adds text-based messages to email """
        part = MIMEText(msg,f'{msgType}')
        
        try:
            if self.status: 
                self.msg.attach(part)
                
        except Exception as e:
            self.add_smtp_error(self.error.get_error_info(e))
            self.status = False
        
        return self
    
    def addAttach(self, file:Path, filename:str) -> 'Mailing':
        """ attach items to email """
        try:
            if self.status:
                with open(file.resolve(), "rb") as attachment:
                    p = MIMEBase('application/pdf', 'octet-stream') 
                    p.set_payload((attachment).read()) 
                    encoders.encode_base64(p) 
                    p.add_header(f'Content-Disposition', f"attachment; filename={filename}") 
                    self.msg.attach(p)
                    
        except Exception as e:
            self.add_smtp_error(self.error.get_error_info(e))
            self.status = False
        
        return self
    
    def addDetails(self, subject:str) -> 'Mailing':
        """ add subject to email """
        try: 
            if self.status: self.msg['Subject'] = subject
        except Exception as e:
            self.add_smtp_error(self.error.get_error_info(e))
            self.status = False
        
        return self
    
    def login(self) -> 'Mailing':
        """ attempts smtp login """
        
        try:
            if self.status: 
                self.smtp = smtplib.SMTP('smtp.gmail.com', 587)
                self.smtp.starttls()
                self.smtp.login(self.email,self.key)
                self.add_smtp_info('Login Successful')
                
        except Exception as e:
            self.add_smtp_error(self.error.get_error_info(e))
            self.status = False
        
        return self
    
    def sendMail(self, toAddr:str) -> 'Mailing':
        """ sends mail """
        try:
            if self.status and self.smtp:
                self.smtp.sendmail(self.email,toAddr,self.msg.as_string())
                self.add_smtp_info(f'Email to ({toAddr}) was send successfully')
        except Exception as e:
            self.add_smtp_error(self.error.get_error_info(e))
            self.status = False
        
        return self
    
    def resetMIME(self) -> 'Mailing':
        """ resets MIMEMultipart() """
        self.msg = MIMEMultipart()
        return self
        
    def destroy(self) -> 'Mailing':
        """ quits current smtp """
        try:
            if self.smtp:
                self.smtp.quit() 
                self.add_smtp_info('Logout Successful')
                
        except Exception as e:
            self.add_smtp_error(self.error.get_error_info(e))
            self.status=False
            
        return self
    
    def add_smtp_error(self, msg:str) -> None:
        self.error.write_error(msg,'SMTP')
        
    def add_smtp_info(self, msg:str) -> None:
        self.error.write_info(msg,'SMTP')


class AsyncMailing:
    """ a class for AsyncMailing (supports method-chaining) """
    def __init__(self, email: str, key: str, error_log: Logger) -> None:
        self.email = email
        self.key = key
        self.error = error_log
        self.status = True
        self.smtp: Optional[aiosmtplib.SMTP] = None

    async def login(self) -> 'AsyncMailing':
        """ Asynchronous SMTP login with persistent connection """
        
        try:
            self.smtp = aiosmtplib.SMTP(hostname="smtp.gmail.com",port=587)
            await self.smtp.connect()
            await self.smtp.login(self.email, self.key)
            self.add_smtp_info("Login Successful")
            
        except Exception as e:
            self.add_smtp_error(self.error.get_error_info(e))
            self.status = False
            
        return self

    async def destroy(self) -> 'AsyncMailing':
        """ Closes the SMTP connection """
        if self.smtp:
            try:
                await self.smtp.quit()
                self.add_smtp_info("Logout Successful")
            except Exception as e:
                self.add_smtp_error(self.error.get_error_info(e))
                self.status = False
        return self
    
        
    async def sendMail(self, toAddr: str, msg:MIMEMultipart) -> bool:
        """ Sends email asynchronously """
        if self.status and self.smtp:
            try:
                msg["From"] = self.email
                await self.smtp.sendmail(self.email, toAddr, msg.as_string())
                self.add_smtp_info(f"Email to ({toAddr}) was sent successfully")
            except Exception as e:
                self.add_smtp_error(self.error.get_error_info(e))
                self.status = False
                
        return self.status

    def add_smtp_error(self, msg: str) -> None:
        self.error.write_error(msg, "ASYNC-SMTP")

    def add_smtp_info(self, msg: str) -> None:
        self.error.write_info(msg, "ASYNC-SMTP")


class AsyncMessage:
    
    def __init__(self, error:Logger):
        self.msg = MIMEMultipart()
        self.error = error
        self.status = True
        
    def addTxtMsg(self, msg: str, msgType: str) -> 'AsyncMessage':
        """ Adds text-based messages to email """
        if self.status:
            try:
                self.msg.attach(MIMEText(msg, msgType))
            except Exception as e:
                self.add_mime_error(self.error.get_error_info(e))
                self.status = False
        return self

    def addAttach(self, file: Path, filename: str) -> 'AsyncMessage':
        """ Attaches files to email """
        if self.status:
            try:
                with open(file.resolve(), "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={filename}")
                    self.msg.attach(part)
            except Exception as e:
                self.add_mime_error(self.error.get_error_info(e))
                self.status = False
        return self

    def addDetails(self, subject: str) -> 'AsyncMessage':
        """ Adds subject to email """
        if self.status:
            try:
                self.msg["Subject"] = subject
            except Exception as e:
                self.add_mime_error(self.error.get_error_info(e))
                self.status = False
        return self
    
    def get_MIME(self) -> MIMEMultipart:
        return self.msg
    
    def add_mime_error(self, msg: str) -> None:
        self.error.write_error(msg, "MIME")