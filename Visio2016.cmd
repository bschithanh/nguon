@echo off
title  Activate Microsoft Office Visio 2016 for FREE - https://github.com/BsChiThanh 
cls
color F4
mode con cols=98 lines=30
 
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

cscript ospp.vbs /inslic:"..\root\Licenses16\pkeyconfig-office.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-bridge-office.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-root.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-root-bridge-test.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-stil.xrm-ms
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-ul.xrm-ms
cscript ospp.vbs /inslic:"..\root\Licenses16\client-issuance-ul-oob.xrm-ms
cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul.xrm-ms"

cscript //nologo ospp.vbs /setprt:1688
cscript //nologo ospp.vbs /sethst:192.168.2.81.2.7.0
cscript //nologo ospp.vbs /sethst:122.226.152.230
cscript //nologo ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK
cscript //nologo ospp.vbs /inpkey:7WHWN-4T7MP-G96JF-G33KR-W8GF4

:end
:notsupported
:halt
pause >nul