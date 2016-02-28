regsvr32 /u "%~dp0IPluginInterface.dll"
%homedrive%
cd %~dp0
ren IPluginInterface.TLB 1.TLB
del IPluginInterface.*
move 1.TLB ..\..\IPluginInterface.TLB