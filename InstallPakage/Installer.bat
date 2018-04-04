@ECHO OFF

REG ADD HKLM\Software\Classes\*\shell\DeliveryUnziper  /d ”[•i•¨‰ð“€
REG ADD HKLM\Software\Classes\*\shell\DeliveryUnziper\command /d "%~dp0Bin\DeliveryUnziper.exe %%1"
