# PdfTableExtraction

# Description
This py script can extract table from pdf with multiple pages contaning clinical data where the keys are same in each table but values differ (will not work on scanned pdfs) and convert into xlsx for easy viewing and other manipulations.

# Installation:
  1. python3.x
  2. XlsxWriter for python3
  3. poppler or poppler-utils containing pdftotext (depending on distro)
  4. OS: GNU/Linux

# Usage:
run PdfTableExtraction.py from the folder containing the PDFs in the example format. For other formats, adjustments needs to be done.
For simplicity, provide the starting and end text of first column.

eg: python3 PdfTableExtraction.py ALL.pdf Disease Note

The result will be a .txt file of same name as pdf. This file helps us to cross check the extraction and to do necessary changes if needed on data change

Second file is .xlsx with same name as pdf. This is the final output.

# License
GPLv3
