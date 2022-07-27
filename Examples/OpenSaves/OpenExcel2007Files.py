import jpype
import asposecells
jpype.startJVM()
from asposecells.api import Workbook, LoadOptions, LoadFormat

# The path to the documents directory.
dataDir = ""

# Opening Microsoft Excel 2007 Xlsx Files
loadOptions2 = LoadOptions(LoadFormat.XLSX)

# Create a Workbook object and opening the file from its path
wbExcel07 = Workbook(dataDir + "Input.xlsx", loadOptions2)
print("Microsoft Excel 2007 - Office365 workbook opened successfully!")

jpype.shutdownJVM()
