@ECHO OFF

REG ADD HKLM\Software\Classes\*\shell\DeliveryUnziper  /d �[�i����
REG ADD HKLM\Software\Classes\*\shell\DeliveryUnziper\command /d "%~dp0Bin\DeliveryUnziper.exe %%1"
