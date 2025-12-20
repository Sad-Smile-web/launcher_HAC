@echo off
echo Установка зависимостей...
python -m pip install pyinstaller pywin32 winshell pythoncom

echo.
echo Проверка наличия архива hac.zip...
if not exist "hac.zip" (
    echo ОШИБКА: Не найден файл hac.zip!
    echo Поместите архив с игрой hac.zip в папку с этим скриптом.
    pause
    exit /b 1
)

echo.
echo Сборка лаунчера со встроенным архивом...
python -m PyInstaller --onefile --windowed --name "HAC_Launcher_Standalone" ^
--add-data "hac.zip;." launcher.py

echo.
echo ========================================
echo Сборка завершена!
echo.
echo Создан единый EXE-файл со встроенной игрой:
echo dist\HAC_Launcher_Standalone.exe
echo.
echo Размер архива hac.zip: 
dir hac.zip | find "hac.zip"
echo ========================================
echo.
pause