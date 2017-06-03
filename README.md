# DocxToPdf
A command line tool which allows to convert DOCX files to PDF.

## Command line args

Command | Alias | Help line
------- | ----- | ---------
--input | -i    | Required. Input .docx file which will be converted.
--output | -o   | Output .pdf file.
--silent | -s   | (Default: False) Remove all log messages from output.
--help | none   | Display this help screen.


## Exemples

Converts _example.docx_ word document into _example.pdf_: `DocxToPdf.exe -i example.docx`

Converts _example.docx_ word document into _file.pdf_: `DocxToPdf.exe -i example.docx -o file.pdf` 

Converts _example.docx_ word document without verbose: `DocxToPdf.exe -i example.docx -s`
