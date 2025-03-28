@echo off
title  Activate Microsoft Office Projet 2019 for FREE - https://github.com/BsChiThanh 
cls
color F4
mode con cols=98 lines=30 
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectPro2019VL_KMS_Client_AE-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectPro2019VL_KMS_Client_AE-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectPro2019VL_KMS_Client_AE-ppd.xrm-ms"
cscript //nologo ospp.vbs /setprt:1688
cscript //nologo ospp.vbs /sethst:192.168.2.81.2.7.0
cscript //nologo ospp.vbs /sethst:122.226.152.230
cscript //nologo ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
:end
:notsupported
:halt
pause >nul
