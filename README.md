# Search Tool
The Search Tool program locates any files containing keywords and compiles them in an
Excel file. It supports word docs, powerpoints, pdfs and will skip password protected files.
The Excel file will contain the file name, file path, and the keywords found. The file path
values are set up as a link for quick file opening.

## Installation
Use the package manager [pip](https://pip.pypa.io/en/stable/) to the required libraries

```bash
pip install openpyxl
pip install comtypes
pip install python-pptx
pip install fitz
pip install aspose
```

## How to use this program

1. First you will be prompted to enter a search path for the folders you wish to search:
○ Either copy the folder path and paste it by right clicking, or manually type in
the folder path
○ After inputting a valid search path (it will check before proceeding), type the
keywords to be searched for (comma-separated).
2. Next, type a name for the output Excel file which will contain the search results
(Default name is “FileSearch”).
3. Next, you will be prompted to enter the destination path for where the Excel file will
be saved. If no path is entered, it will be saved in the same directory as the
application. (e.g. Desktop)
4. Once the search is complete, view the results in the Excel file.
