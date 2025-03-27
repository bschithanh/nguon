@echo off
title  Activate Microsoft Office Prolus 2021 for FREE - https://github.com/BsChiThanh 
cls
color F4
mode con cols=98 lines=30
 
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2021VL_KMS_Client_AE-ppd.xrm-ms
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2021VL_KMS_Client_AE-ul.xrm-ms
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2021VL_KMS_Client_AE-ul-oob.xrm-ms


cscript //nologo ospp.vbs /setprt:1688
cscript //nologo ospp.vbs /sethst:192.168.2.81.2.7.0
cscript //nologo ospp.vbs /sethst:122.226.152.230
cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH

:end
:notsupported
:halt
pause >nul