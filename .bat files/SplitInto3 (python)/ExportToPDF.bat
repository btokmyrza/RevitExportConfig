@echo off
SET mypath=%~dp0
echo %mypath:~0,-1%
echo *******
echo.

python "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Python Scripts\ScanFoldersForRevitFiles_SplitInto3.py" "%mypath:~0,-1%"

%LOCALAPPDATA%\RevitBatchProcessor\BatchRvt.exe --settings_file "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Settings Files\ExportToPDF.json"

python "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Python Scripts\DeleteDummyPagesInPDFs.py" "%mypath:~0,-1%"
