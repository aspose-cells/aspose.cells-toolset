![](https://img.shields.io/badge/REST%20API-v3.0-lightgrey) ![PyPI](https://img.shields.io/pypi/v/excelpandas) ![PyPI - Python Version](https://img.shields.io/pypi/pyversions/spreadsheetpandas) ![PyPI - Downloads](https://img.shields.io/pypi/dm/excelpandas)  [![GitHub license](https://img.shields.io/github/license/aspose-cells/excelpandas)](https://github.com/aspose-cells/excelpandas/blob/main/LICENSE) ![GitHub commits since latest release (by date)](https://img.shields.io/github/commits-since/aspose-cells/SpreadsheetPandas/24.3.0)

# SpreadsheetPandas

SpreadsheetPandas seamlessly merges Aspose.Cells with Pandas, leveraging the robust data processing capabilities of Excel to enhance Pandas' analysis efficiency. Additionally, it utilizes Excel's superior chart display features to present Pandas' analysis results effectively. By harnessing the strengths of both Excel and Pandas, Spreadsheet Pandas offers comprehensive data analysis solutions.

# Why Pandas, Excel, and Aspose.Cells

Aspose.Cells boasts robust spreadsheet processing capabilities, supporting Excel as well as other file formats like ODS, CSV, HTML, and more. It seamlessly converts between these formats, offering a versatile solution for file handling needs.

Excel is a robust spreadsheet software that facilitates data transformation and real-time collaboration. Conversely, Pandas is a potent data processing tool based on Python, widely employed in data science, analysis, and machine learning. Despite Pandas' proficiency in data processing, it lacks certain Excel features. Let's delve into these disparities.
Firstly, Pandas lacks a graphical user interface (GUI) for data manipulation, unlike Excel, which offers intuitive visualization and graphical manipulation. Excel empowers users to sort, filter, and locate data effortlessly through an intuitive interface. Moreover, it furnishes rich charts and graphs for data visualization, aiding users in comprehending data relationships and trends. Although Pandas offers basic drawing functions, its graphical capabilities pale in comparison to Excel.
Secondly, Excel boasts a potent data analysis tool: the pivot table. This feature enables swift data summarization and analysis, allowing users to group, filter, and calculate data for valuable insights. Conversely, Pandas does not natively support pivot tables, necessitating the use of other libraries or tools to achieve similar functions.
Thirdly, Excel offers comprehensive features for cell data such as data validation, comments, and annotations, as well as extensive formatting options. Users can conditionally format, customize fonts, colors, borders, alignment, and more to suit their preferences. In contrast, Pandas prioritizes data processing functions over cell formatting functions.
Fourthly, Excel supports a wide array of formula calculation functions, enabling users to perform complex mathematical, logical, and data analysis operations. Conversely, Pandas lacks Excel's formula calculation functions, relying more on Python syntax and functions for data processing and analysis.
Finally, while Pandas may not offer all of Excel's features, it excels in data processing and analysis. It provides robust functions for data cleansing, transformation, statistical analysis, etc., suitable for various data science and processing tasks. With Pandas, users can efficiently process and analyze data, extracting valuable insights to inform decision-making and forecasting.

# Main Feature

- Read Spreadsheet
- Write Spreadsheet

# Support Formats

|**Format**|**Description**|**Load**|**Save**|
| :- | :- | :- | :- |
|[XLS](https://docs.fileformat.com/spreadsheet/xls/)|Excel 95/5.0 - 2003 Workbook.|&radic;|&radic;|
|[XLSX](https://docs.fileformat.com/spreadsheet/xlsx/)|Office Open XML SpreadsheetML Workbook or template file, with or without macros.|&radic;|&radic;|
|[XLSB](https://docs.fileformat.com/spreadsheet/xlsb/)|Excel Binary Workbook.|&radic;|&radic;|
|[XLSM](https://docs.fileformat.com/spreadsheet/xlsm/)|Excel Macro-Enabled Workbook.|&radic;|&radic;|
|[XLT](https://docs.fileformat.com/spreadsheet/xlt/)|Excel 97 - Excel 2003 Template.|&radic;|&radic;|
|[XLTX](https://docs.fileformat.com/spreadsheet/xltx/)|Excel Template.|&radic;|&radic;|
|[XLTM](https://docs.fileformat.com/spreadsheet/xltm/)|Excel Macro-Enabled Template.|&radic;|&radic;|
|[XLAM](https://docs.fileformat.com/spreadsheet/xlam/)|An Excel Macro-Enabled Add-In file that's used to add new functions to Excel.| |&radic;|
|[CSV](https://docs.fileformat.com/spreadsheet/csv/)|CSV (Comma Separated Value) file.|&radic;|&radic;|
|[TSV](https://docs.fileformat.com/spreadsheet/tsv/)|TSV (Tab-separated values) file.|&radic;|&radic;|
|TabDelimited|Tab-delimited text file, same with TSV file.|&radic;|&radic;|
|[TXT](https://docs.fileformat.com/word-processing/txt/)|Delimited plain text file.|&radic;|&radic;|
|[HTML](https://docs.fileformat.com/web/html/)|HTML format.|&radic;|&radic;|
|[MHTML](https://docs.fileformat.com/web/mhtml/)|MHTML file.|&radic;|&radic;|
|[ODS](https://docs.fileformat.com/spreadsheet/ods/)|ODS (OpenDocument Spreadsheet).|&radic;|&radic;|
|SpreadsheetML|Excel 2003 XML file.|&radic;|&radic;|
|[Numbers](https://docs.fileformat.com/spreadsheet/numbers/)|The document is created by Apple's "Numbers" application which forms part of Apple's iWork office suite, a set of applications which run on the Mac OS X and iOS operating systems.|&radic;||
|[JSON](https://docs.fileformat.com/web/json/)|JavaScript Object Notation|&radic;|&radic;|
|[DIF](https://docs.fileformat.com/spreadsheet/dif/)|Data Interchange Format.| |&radic;|
|[PDF](https://docs.fileformat.com/pdf/)|Adobe Portable Document Format.| |&radic;|
|[XPS](https://docs.fileformat.com/page-description-language/xps/)|XML Paper Specification Format.| |&radic;|
|[SVG](https://docs.fileformat.com/page-description-language/svg/)|Scalable Vector Graphics Format.| |&radic;|
|[TIFF](https://docs.fileformat.com/image/tiff/)|Tagged Image File Format| |&radic;|
|[PNG](https://docs.fileformat.com/image/png/)|Portable Network Graphics Format| |&radic;|
|[BMP](https://docs.fileformat.com/image/bmp/)|Bitmap Image Format| |&radic;|
|[EMF](https://docs.fileformat.com/image/emf/)|Enhanced metafile Format| |&radic;|
|[JPEG](https://docs.fileformat.com/image/jpeg/)|JPEG is a type of image format that is saved using the method of lossy compression.| |&radic;|
|[GIF](https://docs.fileformat.com/image/gif/)|Graphical Interchange Format| |&radic;|
|[MARKDOWN](https://docs.fileformat.com/word-processing/md/)|Represents a markdown document.| |&radic;|
|[SXC](https://docs.fileformat.com/spreadsheet/sxc/)|An XML based format used by OpenOffice and StarOffice|&radic;|&radic;|
|[FODS](https://docs.fileformat.com/spreadsheet/fods/)|This is an Open Document format stored as flat XML.|&radic;|&radic;|
|[DOCX](https://docs.fileformat.com/word-processing/docx/)|A well-known format for Microsoft Word documents that is a combination of XML and binary files.||&radic;|
|[PPTX](https://docs.fileformat.com/presentation/pptx/)|The PPTX format is based on the Microsoft PowerPoint open XML presentation file format.||&radic;|


# Quick Start Guide

To begin with Spreadsheet Pandas, here's what you need to do:
