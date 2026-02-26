@echo off
chcp 936 >nul
cls
echo 脚本运行目录：%~dp0
echo.

:: ========== 第一步：切换到Vue项目根目录（使用相对路径精准跳转） ==========
cd /d ../Frontend
echo 已切换到Vue项目目录：%cd%
echo.

:: ========== 第二步：执行Vue打包命令 生成dist文件夹 ==========
echo 开始执行打包命令：npm run build，请耐心等待...
echo --------------------------------------------------------
npm run build
