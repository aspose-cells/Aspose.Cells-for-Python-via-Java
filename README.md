# Spreadsheet Processing Python High Code API

[Product Page](https://products.aspose.com/cells/python-java/) | [Docs](https://docs.aspose.com/cells/python-java/) | [Demos](https://products.aspose.app/cells/family/) | [API Reference](https://reference.aspose.com//cells/python-java/) | [Examples](https://github.com/aspose-cells/Aspose.Cells-for-Python-via-Java) | [Blog](https://blog.aspose.com/category/cells/) | [Free Support](https://forum.aspose.com/c/cells) | [Temporary License](https://purchase.aspose.com/temporary-license)

[Aspose.Cells for Python via Java](https://products.aspose.com/cells/python-java/) is a scalable and feature-rich API to process Excel&reg; spreadsheets using Python. API offers Excel&reg; file creation, manipulation, conversion and rendering. Developers can format worksheets, rows, columns or cells to the most granular level, create manipulate chart and pivot tables, render worksheets, charts and specific data ranges to PDF or images, add calculate Excel&reg;'s built-in and custom formulas and much more - all without any dependency on Microsoft Office or Excel&reg; application.

## Spreadsheet Python via Java On-premise API Features

- Spreadsheet generation & manipulation via API.
- High-quality file format conversion & rendering.
- Print Microsoft Excel&reg; files to physical or virtual printers.
- Combine, modify, protect, or parse Excel&reg; sheets.
- Apply worksheet formatting.
- Configure and apply page setup for the worksheets.
- Create & customize Excel&reg; charts, Pivot Tables, conditional
  formatting rules, slicers, tables & spark-lines.
- Convert Excel&reg; charts to images & PDF.
- Convert Excel&reg; files to various other formats.
- Formula calculation engine that supports all basic and advanced Excel&reg; functions.

Please visit the [official documentation](https://docs.aspose.com/cells/python-java/) for a more detailed list of features.

## Read & Write Sreadsheet File Formats

**Microsoft Excel&reg;:** XLS, XLSX, XLSB, XLSM, XLT, XLTX, XLTM, CSV, TSV, TabDelimited, SpreadsheetML\
**OpenOffice:** ODS, SXC, FODS\
**Text:** TXT\
**Web:** HTML, MHTML\
**iWork&reg;:** Numbers\
**Other:** SXC, FODS

## Save Spreadsheet Files AS

**Microsoft Word&reg;:** DOCX\
**Microsoft PowerPoint&reg;:** PPTX\
**Microsoft Excel&reg;:** XLAM\
**Fixed Layout:** PDF, XPS\
**Data Interchange:** DIF\
**Vector Graphics:** SVG\
**Image:** TIFF,PNG, BMP, JPEG, GIF\
**Meta File:** EMF\
**Markdown:** MD

Please visit [Supported File Formats](https://docs.aspose.com/cells/python-java/supported-file-formats/) for further details.

## System Requirements

Your machine does not need to have Microsoft Excel&reg; or OpenOffice&reg; software installed.

### Supported Operating Systems

**Microsoft Windows&reg;:** Windows Desktop & Server (`x64`, `x86`)\
**Linux:** Ubuntu, OpenSUSE, CentOS, and others\
**Other:** Any operating system (OS) that can install Java 1.8 or higher, Python 3.5 or higher.

## Get Started

### Installation via `pip`

The Aspose.Cells for Python via Java is [available at pypi.org](https://pypi.org/project/aspose-cells/). To install it, please run the following command:

`pip install aspose-cells`

### Create Excel&reg; File from scratch using Python

```python
#import the python package
import jpype
import asposecells
jpype.startJVM()
from asposecells.api import Workbook

#Create a new Workbook
workbook = Workbook()

#Get the first worksheet
worksheet=workbook.getWorksheets().get(0)

#Get the "A1" cell
cells=worksheet.getCells()
cell=cells.get("A1")

#Write "Hello World" to  "A1" in the first sheet
cell.setValue("Hello World!")

#save this workbook to XLSX 
workbook.save("HelloWorld.xlsx")

jpype.shutdownJVM()
```

## Convert Excel&reg; `XLSX` File to `PDF` using Python

```python
#import the python package
import jpype
import asposecells
jpype.startJVM()
from asposecells.api import Workbook

#Open a existing Workbook
workbook = Workbook("bookwithChart.xlsx")

#save this workbook to PDF file,you can see a chart while open the file with MS Excel&reg;*/
workbook.save("Convert.pdf");

jpype.shutdownJVM()
```

[Product Page](https://products.aspose.com/cells/python-java) | [Docs](https://docs.aspose.com/cells/python-java/) | [Demos](https://products.aspose.app/cells/family/) | [API Reference](https://reference.aspose.com//cells/python-java/) | [Examples](https://github.com/aspose-cells/Aspose.Cells-for-Python-via-Java) | [Blog](https://blog.aspose.com/category/cells/) | [Free Support](https://forum.aspose.com/c/cells) | [Temporary License](https://purchase.aspose.com/temporary-license)
