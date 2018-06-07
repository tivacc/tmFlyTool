@echo off
cd commander
commander.exe  device lock --debug disable --device EFR32
echo 按任意键退出 & pause
exit