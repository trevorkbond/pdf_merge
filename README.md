# Usage and Purpose
This PDF merger is specifically designed for use at a private, professional firm. The firm manages and compiles source documentation for clients with a specific naming convention and creates
"scans" of their work. This merger was created because naming conventions consistently align with where pdf's should be placed within the scan and can therefore be compiled in a way that
code can understand.

# Requirements
Usage of this requires installation of PyPDF. However, the finalized .app version (created using py2app) runs on current MacOS without any updating of Python or installation of PyPDF.

# Instructions
Upon execution of the code, a TkInter UI will appear prompting input of the scan folder (where all source docs are located), the scan template file name, and the desired name of the outputted
merged scan. The code searches for substrings among the names of the pdf files within the designated scan folder. See the following for a breakdown of the key substrings and the
section they will be merged into:

- "Overview" - Overview
- "Financials" - Period Financials (Balance Sheet, P&L, P&L Year to Date)
- "REC B" & "REC CC" - Reconciliations
- "STMT B" & "STMT CC" - Bank and Credit Card Statements
- "AJE R" - Adjusting Journal Entries
- "OI AJE" - Adjusting Journal Entry source documents
- "PR" - Payroll Reports
- "DEP" - Depreciation Schedule
- "STMT L" - Loan Statements
- "SCH L" - Retained Earnings
- "OI" - Other Source Documents
- "CLIENT COM" - Client Communications
See the [example scan template](https://github.com/trevorkbond/pdf_merge/blob/main/Template.pdf) for more information on file ordering.

Upon successful location of the scan folder and merging of files, the merged pdf will be output to the user's desktop with the given filename. All documents not named with the proper conventions,
along with all non-pdf files, will be excluded from the scan.
