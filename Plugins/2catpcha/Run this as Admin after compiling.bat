set plugin=2captcha

ren %plugin%.dll 1.dll
del %plugin%.*
ren 1.dll %plugin%.dll
if not exist %windir%\SysWOW64 goto nosyswow
%homedrive%
cd %windir%\SysWOW64
:nosyswow
regsvr32 /u "%~dp0%plugin%.dll"
%~d0
cd %~p0
move %plugin%.dll ..