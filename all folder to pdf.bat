@echo off

REM Get the folder path of the batch file
set "folder=%~dp0"

REM Execute the PDF command on all .docx files within the folder
for %%f in ("%folder%\*.docx") do (
    SBCCmd -d "%%f" -o "%%~nf.pdf"
)