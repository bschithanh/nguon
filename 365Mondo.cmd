@echo off
title  Activate Microsoft Mondo 2016 for FREE - https://github.com/BsChiThanh 
cls
color F4
mode con cols=98 lines=30
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
cscript ospp.vbs /inslic:"..\root\Licenses16\MondoVL_MAK-ul-phn.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\MondoVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\MondoVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\MondoVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\MondoVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\MondoVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\MondoVL_KMS_Client-ppd.xrm-ms"
cscript //nologo slmgr.vbs /ckms
cscript //nologo ospp.vbs /setprt:1688
cscript //nologo ospp.vbs /sethst:192.168.2.81.2.7.0
cscript //nologo ospp.vbs /inpkey:HFTND-W9MK4-8B7MJ-B6C4G-XQBR2
:end
:notsupported
:halt
pause >nul
