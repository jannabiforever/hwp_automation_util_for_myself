import win32com.client as win32
from abc import ABC, abstractmethod
import logging


CLSID = "HWPFrame.HwpObject"


class ServiceLogger:
    def __init__(self, sv_name: str):
        self.sv_name = sv_name

    def log(self, func):
        def wrapper(*args, **kwargs):
            try:
                logging.info(f"Executing {self.sv_name}")
                result = func(*args, **kwargs)
                logging.info(
                    f"Result of {self.sv_name}: {result}")
            except Exception as e:
                logging.error(
                    f"Error in {self.sv_name}: {e}")
                raise e

        return wrapper


class HApp:
    __hwp_com_obj = None

    def __new__(cls):
        if cls.__hwp_com_obj is None:
            cls.__hwp_com_obj = win32.Dispatch(CLSID)
            # To open hwp application without any window for extra permissions,
            # we need to register predefined module.
            cls.__hwp_com_obj.RegisterModule(
                "FilePathCheckDLL", "FilePathCheckerModule")
            cls.__hwp_com_obj.XHwpWindows.Item(0).Visible = True

            return cls.__hwp_com_obj

        return cls.__hwp_com_obj


class HService(ABC):
    @abstractmethod
    def execute(self, com) -> tuple:
        """
        Execute a certain service using the HWP application object.

        Every service should return a tuple, even though it is empty.
        """
        raise NotImplementedError("execute method must be implemented")


class HServiceBuilder(ABC):
    @abstractmethod
    def build(self) -> HService:
        """
        Build a service object.
        """
        pass
