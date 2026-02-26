@echo off
chcp 936 >nul
cls

echo          Vue打包文件 复制到后端wwwroot

:: 强制切换到Vue目录，绝对生效
cd /d "%~dp0../Frontend"
echo 当前Vue项目目录：%cd%
echo.

:: 检查dist文件夹是否存在+是否为空
if not exist "dist" (
    echo 【错误】未找到 dist 文件夹，请先执行打包命令！
    pause
    exit /b 1
)

echo 【正常】检测到dist文件夹，内含文件，准备复制

:: 仅清空wwwroot下的dist文件夹（核心修改处）
echo 正在清空 wwwroot 下的 dist 旧文件...
rd /s /q "%~dp0../Backend/CenterBackend/wwwroot/dist" 2>nul
:: 无需重建整个wwwroot，仅在删除dist后（如需）确保dist目录存在（xcopy也会自动创建）
md "%~dp0../Backend/CenterBackend/wwwroot/dist" 2>nul
echo 清空完成！
echo.

echo 正在复制dist文件到wwwroot目录...
xcopy "dist" "%~dp0../Backend/CenterBackend/wwwroot/dist" /s /e /y /q /i
echo 复制完成！
echo.

echo ========================================================
echo 复制成功！无任何异常！
echo 源文件夹：%cd%\dist
echo 目标文件夹：%~dp0Backend\CenterBackend\wwwroot\dist
echo ========================================================
echo.
pause