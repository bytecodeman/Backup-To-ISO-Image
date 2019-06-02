@ECHO OFF

SET hr=%time:~0,2%
if "%hr:~0,1%" equ " " SET hr=0%hr:~1,1%

robocopy c:\sqlserverbackup "C:\Users\Administrator\Google Drive\sqlbackups" /MIR /FFT /R:3 /W:10 /Z /NP /LOG:c:\temp\rcstat_sqlbackup_%date:~-4,4%%date:~-10,2%%date:~-7,2%_%hr%%time:~3,2%%time:~6,2%.txt

SET hr=%time:~0,2%
if "%hr:~0,1%" equ " " SET hr=0%hr:~1,1%

robocopy c:\inetpub\wwwroot "C:\Users\Administrator\Google Drive\wwwroot" /MIR /FFT /R:3 /W:10 /Z /NP /LOG:c:\temp\rcstat_www_%date:~-4,4%%date:~-10,2%%date:~-7,2%_%hr%%time:~3,2%%time:~6,2%.txt
