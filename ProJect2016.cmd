@echo off
title  Activate Microsoft Office Projet 2016 for FREE - https://github.com/BsChiThanh 
cls
color F4
mode con cols=98 lines=30
 
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-root.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-stil.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-ul-oob.xrm-ms"


cscript //nologo ospp.vbs /setprt:1688
cscript //nologo ospp.vbs /sethst:192.168.2.81.2.7.0
cscript //nologo ospp.vbs /sethst:122.226.152.230
cscript //nologo ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT
cscript //nologo ospp.vbs /inpkey:GNFHQ-F6YQM-KQDGJ-327XX-KQBVC

:end
:notsupported
:halt
pause >nul