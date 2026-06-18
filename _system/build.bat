@echo off
setlocal

echo ============================================================
echo   Budowanie Importuj.exe / Raport.exe / Ekstrakt.exe / Wyjatki.exe
echo ============================================================

echo.
echo [1/6] Instalowanie zaleznosci...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo [2/6] Czyszczenie build\ i dist\...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul

echo.
echo [3/6] Budowanie Importuj.exe...
pyinstaller --onefile --clean --name "Importuj" importuj.py

echo.
echo [4/6] Budowanie Raport.exe...
pyinstaller --onefile --clean --name "Raport" --collect-all openpyxl raport.py

echo.
echo [5/6] Budowanie Ekstrakt.exe...
pyinstaller --onefile --clean --name "Ekstrakt" --collect-all openpyxl ekstrakt.py

echo.
echo [6/6] Budowanie Wyjatki.exe...
pyinstaller --onefile --clean --name "Wyjatki" wyjatki.py

echo.
echo Kopiowanie .exe do _system\ i do folderu glownego...
copy /Y dist\Importuj.exe .
copy /Y dist\Raport.exe .
copy /Y dist\Ekstrakt.exe .
copy /Y dist\Wyjatki.exe .
copy /Y dist\Importuj.exe ..\
copy /Y dist\Raport.exe ..\
copy /Y dist\Ekstrakt.exe ..\
copy /Y dist\Wyjatki.exe ..\

echo.
echo ============================================================
echo   Gotowe! Nowe .exe sa w _system\ i w folderze glownym.
echo ============================================================

endlocal
