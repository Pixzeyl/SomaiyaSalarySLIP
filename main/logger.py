import logging
from pathlib import Path

class Logger:
    def __init__(self, dir_path: Path) -> None:
        self.log = logging.getLogger(__name__)
        file_path = Path(dir_path,'app.log')
        console, file = logging.StreamHandler(), logging.FileHandler(file_path,mode="a",encoding="utf-8")
        formatter = logging.Formatter("[{asctime}]:[{levelname}]:{message}",style="{",datefmt="%d-%m-%Y %H:%M")
        console.setFormatter(formatter)
        file.setFormatter(formatter)
        self.log.addHandler(console)        
        self.log.addHandler(file)        
        self.log.setLevel(logging.DEBUG)
                
    def write_error(self, msg:str, where:str = 'APPLICATION') -> None:
        self.log.warning(msg=f"[{where}] {msg}",exc_info=True)

    def write_info(self, msg:str, where:str = 'APPLICATION') -> None:
        self.log.info(f"[{where}] {msg}")
    
    def get_error_info(self, exception:Exception) -> str:
        if(isinstance(exception,Exception)):
            return f"{exception.__class__.__module__}.{exception.__class__.__name__} : {exception}"
        else:
            return "Not an exception"