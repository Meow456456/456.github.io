@echo off
cscript c:\windows\system32\slmgr.vbs -skms kms.pu.edu.tw:1688
cscript c:\windows\system32\slmgr.vbs -ato
cscript c:\windows\system32\slmgr.vbs -dli

@echo off
echo 本程式將執行Office Professional Plus 2019版本的啟動認證的設定
echo 請確認您執行本程式時是否已經選擇以***系統管理者身分執行***
@echo 此認證程式只有在 靜宜大學 所授權範圍使用有效

echo 偵測 Microsoft Office 2019 安裝目錄
set ProgRoot=%ProgramFiles%
if not "%ProgramFiles%" == "" set ProgRoot=%ProgramFiles%
echo 設定 KMS 指向 靜宜大學 金鑰管理伺服器 - kms.pu.edu.tw
cscript "%ProgRoot%\Microsoft Office\Office16\ospp.vbs" /sethst:kms.pu.edu.tw
cscript "%ProgRoot%\Microsoft Office\Office16\ospp.vbs" /setprt:1688
echo 啟動 Microsoft Office 2019 
cscript "%ProgRoot%\Microsoft Office\Office16\ospp.vbs" /act
echo 啟動程序執行完成
echo 請注意: 上方(約前五行)需有 Product activation successful 出現, 
echo 才表示您的 Office2019 啟動成功!
exit