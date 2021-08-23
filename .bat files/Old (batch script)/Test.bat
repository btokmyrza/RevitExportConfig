@echo off
setlocal ENABLEDELAYEDEXPANSION
chcp 65001
set word=\\iaspserv\iasp\
set str=I:\17_ОРЫНБОР 4.5\1_PROJECT\ORN1_0_BIM коллизии\BIM.rvt
set replace=I:\
set str=%str:I:\=!word!%

echo %str%

endlocal
pause