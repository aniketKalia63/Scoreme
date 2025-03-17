# Scoreme

PDF to Excel Table Extractor
A C# console app that extracts tables from PDF files and saves them to Excel.

Features
Extracts tables based on word positions.
Auto-detects columns.
Exports to Excel using EPPlus.

Dependencies
UglyToad.PdfPig
EPPlus

Usage
Set file paths:
1.
string pdfPath = @"C:\path\to\input.pdf";
string excelPath = @"C:\path\to\output.xlsx";
2.
Run the program.

Notes
Works best with well-structured PDFs.
May need tweaking for complex tables.
