@echo off
rem Builds "Invoice Creator.exe" (Python + the app + the template in one file).
rem Only needed on the computer that makes the .exe - the .exe itself needs nothing installed.
cd /d "%~dp0"
python -m pip install --quiet --disable-pip-version-check "pyinstaller==6.22.3" || goto :failed
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
    --name "Invoice Creator" ^
    --icon "%~dp0app.ico" ^
    --add-data "%~dp0invoice_template.docx;." ^
    --distpath "%~dp0." ^
    --workpath "%TEMP%\invoice-creator-build" ^
    --specpath "%TEMP%\invoice-creator-build" ^
    invoice_generator.py || goto :failed
echo.
echo   Done: "Invoice Creator.exe" is in this folder.
pause
exit /b

:failed
echo.
echo   Build failed - see the messages above.
pause
