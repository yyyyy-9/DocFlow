@echo off
setlocal
chcp 65001 >nul
echo 清理已完成三表任务的中间数据，先保留并核对各阶段总稿。
echo 仅留总稿模式会结束旧02/03接续；可传 --keep-latest 保留最新修订链。
echo 不带路径时处理本包outputs；可把任务父目录拖到此文件上。
call "%~dp0run.cmd" cleanup %* --apply
set "WPS_CLEANUP_EXIT=%errorlevel%"
echo.
if not "%WPS_CLEANUP_EXIT%"=="0" echo 部分操作未完成，请查看上方原因；已保存总稿不会覆盖。
pause
exit /b %WPS_CLEANUP_EXIT%
