ren Math.dll 1.dll
del Math.*
ren 1.dll Math.dll
regsvr32 /u "%~dp0Math.dll"
move Math.dll ..