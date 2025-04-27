
import win32com.client as win32
import os
import pythoncom


pythoncom.CoInitialize()
HWPX = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')  # type: ignore
HWPX.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")


def generate_pdfs(files):
    pdfs = []
    for file in files:
        pdfs.append(generate_pdf(file))
    return pdfs


def generate_pdf(file):
    try:
        HWPX.Open(file)

        dest_path = os.path.splitext(file)[0] + '.pdf'

        param_set = HWPX.HParameterSet.HFileOpenSave

        HWPX.HAction.GetDefault("FileSaveAsPdf", param_set.HSet)

        param_set.filename = dest_path
        param_set.Format = "PDF"

        HWPX.HAction.Execute("FileSaveAsPdf", param_set.HSet)

    finally:
        HWPX.XHwpDocuments.Close(False)
        return os.path.splitext(file)[0] + '.pdf'
