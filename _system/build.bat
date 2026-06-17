@echo off
chcp 65001 >nul
echo ============================================================
echo  BUDOWANIE APLIKACJI — LYRECO FAKTURY
echo ============================================================
echo.

echo [1/4] Instalowanie zaleznosci...
pip install -r requirements.txt
if errorlevel 1 (
    echo BLAD: pip install nie powiodl sie.
    pause & exit /b 1
)

echo.
echo [2/4] Czyszczenie poprzednich buildow...
if exist build  rmdir /s /q build
if exist dist   rmdir /s /q dist
if exist Importuj.spec del /q Importuj.spec
if exist Raport.spec   del /q Raport.spec

echo.
echo [3/4] Budowanie Importuj.exe...
pyinstaller --onefile --clean --name "Importuj" --collect-all openpyxl importuj.py
if errorlevel 1 (
    echo BLAD przy budowaniu Importuj.exe
    pause & exit /b 1
)

echo.
echo [4/5] Budowanie Raport.exe...
pyinstaller --onefile --clean --name "Raport" --collect-all openpyxl raport.py
if errorlevel 1 (
    echo BLAD przy budowaniu Raport.exe
    pause & exit /b 1
)

echo.
echo [5/5] Budowanie Ekstrakt.exe...
pyinstaller --onefile --clean --name "Ekstrakt" --collect-all openpyxl ekstrakt.py
if errorlevel 1 (
    echo BLAD przy budowaniu Ekstrakt.exe
    pause & exit /b 1
)

echo.
echo Kopiowanie plikow exe do biezacego folderu (_system)...
copy /Y dist\Importuj.exe .
copy /Y dist\Raport.exe .
copy /Y dist\Ekstrakt.exe .

echo.
echo Kopiowanie plikow exe do folderu glownego...
copy /Y dist\Importuj.exe ..\
copy /Y dist\Raport.exe ..\
copy /Y dist\Ekstrakt.exe ..\

echo.
echo ============================================================
echo  GOTOWE!
echo  Pliki w _system\:    Importuj.exe, Raport.exe, Ekstrakt.exe
echo  Pliki w folderze gl.: Importuj.exe, Raport.exe, Ekstrakt.exe
echo ============================================================
pause
