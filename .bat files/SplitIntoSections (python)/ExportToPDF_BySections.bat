@echo off
SET mypath=%~dp0
echo %mypath:~0,-1%
echo *******
echo.

python "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Python Scripts\ScanFoldersForRevitFiles_SplitIntoSections.py" "%mypath:~0,-1%"

%LOCALAPPDATA%\RevitBatchProcessor\BatchRvt.exe --settings_file "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Settings Files\BySections\ExportToPDF_AR.json"
%LOCALAPPDATA%\RevitBatchProcessor\BatchRvt.exe --settings_file "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Settings Files\BySections\ExportToPDF_EM.json"
%LOCALAPPDATA%\RevitBatchProcessor\BatchRvt.exe --settings_file "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Settings Files\BySections\ExportToPDF_KG.json"
%LOCALAPPDATA%\RevitBatchProcessor\BatchRvt.exe --settings_file "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Settings Files\BySections\ExportToPDF_OV.json"
%LOCALAPPDATA%\RevitBatchProcessor\BatchRvt.exe --settings_file "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Settings Files\BySections\ExportToPDF_SH.json"
%LOCALAPPDATA%\RevitBatchProcessor\BatchRvt.exe --settings_file "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Settings Files\BySections\ExportToPDF_SPS.json"
%LOCALAPPDATA%\RevitBatchProcessor\BatchRvt.exe --settings_file "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Settings Files\BySections\ExportToPDF_SS.json"
%LOCALAPPDATA%\RevitBatchProcessor\BatchRvt.exe --settings_file "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Settings Files\BySections\ExportToPDF_VK.json"

python "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Python Scripts\DeleteDummyPagesInPDFs.py" "%mypath:~0,-1%"
