@echo off
cd /d "%~dp0"

rem The packaged program needs no Python at all.
if exist "Invoice Creator.exe" (start "" "Invoice Creator.exe" & exit /b)

rem Otherwise find Python: the "py" launcher works even when Python was not added to PATH.
where pyw >nul 2>nul && (start "" pyw invoice_generator.py & exit /b)
where pythonw >nul 2>nul && (start "" pythonw invoice_generator.py & exit /b)
for /d %%D in ("%LocalAppData%\Programs\Python\Python3*" "%ProgramFiles%\Python3*" "%SystemDrive%\Python3*") do (
    if exist "%%~D\pythonw.exe" (start "" "%%~D\pythonw.exe" invoice_generator.py & exit /b)
)

echo.
echo   Python was not found on this computer.
echo.
echo   Easiest fix: put "Invoice Creator.exe" in this folder - it needs nothing installed.
echo   Or install Python from https://www.python.org/downloads/
echo   and tick "Add python.exe to PATH" during the install.
echo.
pause
