@echo off
chcp 1251
echo "--- �������� ��������(�����) � ������! ---"
start /wait "" "C:\Program Files (x86)\iCopy\iCopy.exe\" /p /f /path "%~dp0��������.jpg" /quality 85
echo "--- �������� ���� � ������" ---"
start /wait "" "C:\Program Files (x86)\iCopy\iCopy.exe\" /p /f /path "%~dp0����.jpg" /quality 85
echo "--- ��������� okna.pdf �� stroika@megastroy-5.ru? ---"
..\..\�������\script\curl.exe -L "https://docs.google.com/spreadsheets/d/1T1Gn-XeexJsnbZ4C6XOHQ_8D637JsoyoEFi84Xjkao4/export?exportFormat=pdf&gid=0&range=A:H&top_margin=0.25&bottom_margin=0.25&left_margin=0.3&right_margin=0.2" -o okna.pdf
pause
..\..\�������\script\mailsend1.19.exe -t stroika@megastroy-5.ru -sub "OKHA" -cs Windows-1251 -attach "okna.pdf" -smtp smtp.mail.ru -port 465 -f oknanazakaz.ms@mail.ru -name "oknanazakaz.ms@mail.ru" -rt oknanazakaz.ms@mail.ru -ssl -auth-login -user oknanazakaz.ms@mail.ru -pass account.mail.ru -q && echo "--- ������� ���������� ---" && del okna.pdf
echo "--- ������� ����������� ����? ---"
pause
del %0