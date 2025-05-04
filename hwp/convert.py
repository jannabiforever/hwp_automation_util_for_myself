from .app import HApp, HService, HServiceBuilder, ServiceLogger


class ConvertToPdfServiceBuilder(HServiceBuilder):
    def __init__(self):
        self.file_path = None
        self.save_path = None

    def set_file_path(self, file_path: str) -> "ConvertToPdfServiceBuilder":
        self.file_path = file_path
        return self

    def set_save_path(self, save_path: str) -> "ConvertToPdfServiceBuilder":
        if save_path is not None and not save_path.endswith(".pdf"):
            raise ValueError("save_path must end with .pdf")
        self.save_path = save_path
        return self

    def build(self) -> HService:
        if self.file_path is None:
            raise ValueError("file_path must be set")
        return ConvertToPdfService(self.file_path, self.save_path)


class ConvertToPdfService(HService):
    def __init__(self, file_path: str, save_path: str | None):
        self.file_path = file_path
        if save_path is None:
            save_path = file_path.rsplit(".", 1)[0] + ".pdf"
        elif not save_path.endswith(".pdf"):
            raise ValueError("save_path must end with .pdf")
        self.save_path = save_path

    @ServiceLogger("ConvertToPdfService").log
    def execute(self, com) -> tuple:
        """
        returns:
            - file_path: The path of the original HWP file.
            - save_path: The path where the PDF file is saved.
        """
        com.Open(self.file_path)

        param_set = com.HParameterSet.HFileOpenSave
        com.HAction.GetDefault("FileSaveAsPdf", param_set.HSet)
        param_set.filename = self.save_path
        param_set.Format = "PDF"
        com.HAction.Execute("FileSaveAsPdf", param_set.HSet)

        return (self.file_path, self.save_path)
