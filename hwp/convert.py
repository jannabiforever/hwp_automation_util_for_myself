from .app import HService, HServiceBuilder, ServiceLogger
from .chain import chain
import os
from typing import List


class ConvertToPdfServiceBuilder(HServiceBuilder):
    def __init__(self):
        self.target_path = None
        self.save_path = None

    def set_file_path(self, target_path: str | List[str]) -> "ConvertToPdfServiceBuilder":
        self.target_path = target_path
        return self

    def set_save_path(self, save_path: str | List[str]) -> "ConvertToPdfServiceBuilder":
        if isinstance(save_path, list):
            if not isinstance(self.target_path, list):
                raise ValueError(
                    "save_path must be a list if target_path is a list")
            if any(not p.endswith(".pdf") for p in save_path):
                raise ValueError("save_path must end with .pdf")
        elif isinstance(save_path, str) and not save_path.endswith(".pdf"):
            raise ValueError("save_path must end with .pdf")

        self.save_path = save_path
        return self

    def build(self) -> HService:
        if self.target_path is None:
            raise ValueError("target_path must be set")

        if isinstance(self.target_path, list):
            # multiple hwp -> multiple pdf
            if self.save_path is None:
                # if none, just convert to pdf with same name.
                self.save_path = [
                    p.rsplit(".", 1)[0] + ".pdf" for p in self.target_path]
            elif isinstance(self.save_path, str):
                # In case save_path is a directory
                if not os.path.isdir(self.save_path):
                    raise ValueError(
                        "Given save_path is str, target_path is a list. So save_path must be a directory.")

                stems = [os.path.splitext(os.path.basename(p))[0]
                         for p in self.target_path]
                self.save_path = [os.path.join(
                    self.save_path, f"{s}.pdf") for s in stems]

            return chain(*[ConvertToPdfService(t, s) for t, s in zip(self.target_path, self.save_path)])

        else:
            # single hwp -> single pdf
            # ensured that target_path is str.
            if self.save_path is None:
                self.save_path = self.target_path.rsplit(".", 1)[0] + ".pdf"

            return ConvertToPdfService(self.target_path, self.save_path)


class ConvertToPdfService(HService):
    def __init__(self, target_path: str, save_path: str):
        self.target_path = target_path
        if not save_path.endswith(".pdf"):
            raise ValueError("save_path must end with .pdf")
        self.save_path = save_path

    @ServiceLogger("ConvertToPdfService").log
    def execute(self, com) -> tuple:
        """
        returns:
            - target_path: The path of the original HWP file.
            - save_path: The path where the PDF file is saved.
        """
        com.Open(self.target_path)

        param_set = com.HParameterSet.HFileOpenSave
        com.HAction.GetDefault("FileSaveAsPdf", param_set.HSet)
        param_set.filename = self.save_path
        param_set.Format = "PDF"
        com.HAction.Execute("FileSaveAsPdf", param_set.HSet)

        return (self.target_path, self.save_path)


class ConvertToHwpxServiceBuilder(HServiceBuilder):
    def __init__(self):
        self.target_path = None
        self.save_path = None

    def set_file_path(self, target_path: str | List[str]) -> "ConvertToPdfServiceBuilder":
        self.target_path = target_path
        return self

    def set_save_path(self, save_path: str | List[str]) -> "ConvertToPdfServiceBuilder":
        if isinstance(save_path, list):
            if not isinstance(self.target_path, list):
                raise ValueError(
                    "save_path must be a list if target_path is a list")
            if any(not p.endswith(".pdf") for p in save_path):
                raise ValueError("save_path must end with .pdf")
        elif isinstance(save_path, str) and not save_path.endswith(".pdf"):
            raise ValueError("save_path must end with .pdf")

        self.save_path = save_path
        return self

    def build(self) -> HService:
        if self.target_path is None:
            raise ValueError("target_path must be set")

        if isinstance(self.target_path, list):
            # multiple hwp -> multiple pdf
            if self.save_path is None:
                # if none, just convert to pdf with same name.
                self.save_path = [
                    p.rsplit(".", 1)[0] + ".pdf" for p in self.target_path]
            elif isinstance(self.save_path, str):
                # In case save_path is a directory
                if not os.path.isdir(self.save_path):
                    raise ValueError(
                        "Given save_path is str, target_path is a list. So save_path must be a directory.")

                stems = [os.path.splitext(os.path.basename(p))[0]
                         for p in self.target_path]
                self.save_path = [os.path.join(
                    self.save_path, f"{s}.pdf") for s in stems]

            return chain(*[ConvertToHwpxService(t, s) for t, s in zip(self.target_path, self.save_path)])

        else:
            # single hwp -> single pdf
            # ensured that target_path is str.
            if self.save_path is None:
                self.save_path = self.target_path.rsplit(".", 1)[0] + ".pdf"

            return ConvertToHwpxService(self.target_path, self.save_path)


class ConvertToHwpxService(HService):
    def __init__(self, target_path: str, save_path: str | None):
        self.target_path = target_path
        if save_path is None:
            save_path = target_path.rsplit(".", 1)[0] + ".hwpx"
        elif not save_path.endswith(".hwpx"):
            raise ValueError("save_path must end with .hwpx")
        self.save_path = save_path

    @ServiceLogger("ConvertToHwpxService").log
    def execute(self, com) -> tuple:
        """
        returns:
            - target_path: The path of the original HWP file.
            - save_path: The path where the HWPX file is saved.
        """
        com.Open(self.target_path)

        param_set = com.HParameterSet.HFileOpenSave
        com.HAction.GetDefault("FileSaveAs_S", param_set.HSet)
        param_set.filename = self.save_path
        param_set.Format = "HWPX"
        com.HAction.Execute("FileSaveAs_S", param_set.HSet)

        return (self.target_path, self.save_path)
