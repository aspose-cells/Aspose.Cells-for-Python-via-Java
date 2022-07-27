import jpype
import asposecells
jpype.startJVM()
from asposecells.api import Workbook
from jpype import java

fis = java.io.FileInputStream("Input.xlsx")
workbook = Workbook(fis)
print("Workbook opened using stream successfully!!")
workbook.save("Output.pdf")
fis.close()

jpype.shutdownJVM()
