@echo off
chcp 65001 >nul
echo ============================================
echo   启动 OpenVINO Workshop
echo ============================================
echo.

rem 激活虚拟环境
call ov_workshop\Scripts\activate.bat

rem 启动 JupyterLab
echo 正在启动 JupyterLab ...
jupyter lab .

rem 退出
call deactivate
pause
