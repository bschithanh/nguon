@echo off
title  Activate Microsoft Office Prolus 2016 - https://github.com/BsChiThanh 
cls
cls
color F4
mode con cols=98 lines=30
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms"  
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-phn.xrm-ms
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-pl.xrm-ms
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ppd.xrm-ms
cscript //nologo ospp.vbs /setprt:1688
cscript //nologo ospp.vbs /sethst:192.168.2.81.2.7.0
cscript //nologo ospp.vbs /sethst:122.226.152.230
cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 
cscript //nologo ospp.vbs /inpkey:74NFR-DDJ8W-JBM6M-MKMHC-FX83Y
:end
:notsupported
:halt
pause >nul
