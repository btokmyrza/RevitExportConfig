@echo off
setlocal enabledelayedexpansion
SET mypath=%~dp0
echo %mypath:~0,-1%
echo *******
echo.

break>"\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\revit_file_list.txt"
Set "out=\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\"

for /R %%f in (*.rvt) do (

	echo %%f | findstr /i /c:"1_Design" >nul
	if errorlevel 1 (
		rem echo Folder named "1_Design" not found!

		echo %%f | findstr /i /c:"1_PROJECT" >nul
		if errorlevel 1 (
			rem echo Folder named "1_PROJECT" not found!
		) else (
			for /f "delims=#" %%a in ("%%f") do (
				if "%%a"=="%%f" (
					echo %%f >> "%out%\revit_file_list.txt"
				) else (
					echo It contains #
					echo %%f
					echo -------
				)
			)
		)

	) else (
		for /f "delims=#" %%a in ("%%f") do (
			if "%%a"=="%%f" (
				echo %%f >> "%out%\revit_file_list.txt"
			) else (
				echo It contains #
				echo %%f
				echo -------
			)
		)
	)

)

%LOCALAPPDATA%\RevitBatchProcessor\BatchRvt.exe --settings_file "\\iasp-dc2\Documents\05_Soft\Revit Batch Processor\RevitExportConfig\Settings Files\ExportToDWF.json"

endlocal
pause
