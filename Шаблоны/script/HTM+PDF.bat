@echo off
chcp 65001

rem Инструкция: 
rem Заполнить в СуперОкна7 поле КЛИЕНТ: 
rem Клиент: ТРИ слова разделённые пробелами 
rem Адрес: Строка от 2 до 44 символов 
rem Телефон: 11 цифр, перед которыми стоит знак + 

rem Нажать кнопку: 
rem Печать > Просмотр > Сохранить > Выберите папку КП! > Файл HTM > 
rem Имя файла будет использовано как примечание при создании папки 
rem > Сохранить 

rem Перетащите созданный файл HTM на файл HTM+PDF.bat находящийся в той же папке


if not "%~x1" == ".htm" if not "%~x2" == ".htm" (echo --- Затащите на этот файл мышкой, файлы HTM и PDF --- && pause && exit)
set "dp0=%~dp0"
FOR /F "tokens=*" %%a IN ('type %dp0:~0,-1%:user.drive.id') DO set "gfid=%%a"
IF "%~x1" == ".pdf" (start /wait "" "C:\Program Files\SumatraPDF\SumatraPDF.exe" -print-to "HP_M132nw" -print-settings "3x" %~nx1)
IF "%~x2" == ".pdf" (start /wait "" "C:\Program Files\SumatraPDF\SumatraPDF.exe" -print-to "HP_M132nw" -print-settings "3x" %~nx2) 
echo --- Выполняется запрос doGet ---
FOR /F "tokens=*" %%a IN ('..\Шаблоны\script\curl.exe -G --data-urlencode "fileName1=%~nx1" --data-urlencode "fileName2=%~nx2" --data-urlencode "folderId=%gfid%" -L "https://script.google.com/macros/s/AKfycbxYDA4ngMOHdLyiBpS3jegwI1Q4Ox8I84dvk1EX/exec"') DO set "response=%%a"
for /F "tokens=1,2 delims=;" %%a in ("%response%") do (
   start /wait "Скачивание сформированного договора" ..\Шаблоны\script\curl.exe -L "https://docs.google.com/document/d/%%b/export?exportFormat=pdf" -o Договор.pdf
   echo --- Распечатка первой страницы договора --- 
   start /wait "" "C:\Program Files\SumatraPDF\SumatraPDF.exe" -print-to "HP_M132nw" -print-settings "1,3x" Договор.pdf
   rem start /wait "" "C:\Program Files\SumatraPDF\SumatraPDF.exe" -page 1 -print-dialog -exit-when-done Договор.pdf
   echo --- Переверните распечатанную страницу договора --- 
   pause
   start /wait "" "C:\Program Files\SumatraPDF\SumatraPDF.exe" -print-to "HP_M132nw" -print-settings "2,3x" Договор.pdf
   echo --- После закрытия этого окна запустите файл СКАНИРОВАТЬ.bat ---
   timeout 20
   rem start /wait "" "C:\Program Files\SumatraPDF\SumatraPDF.exe" -page 2 -print-dialog -exit-when-done Договор.pdf
   move Договор.pdf "%%a"
   explorer /select, "%%a\СКАНИРОВАТЬ.bat"
)
