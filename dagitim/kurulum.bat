@echo off
setlocal enabledelayedexpansion
title Nobet Planlama Sistemi - Kurulum

set "EXENAME=NobetPlanlamaSistemi.exe"
set "SRC=%~dp0%EXENAME%"
set "APPDIR=%LOCALAPPDATA%\NobetPlanlamaSistemi"
set "PS1=%TEMP%\nobet_kurulum_%RANDOM%.ps1"

echo ================================================
echo   Nobet Planlama Sistemi - Kurulum
echo ================================================
echo.

if not exist "%SRC%" (
    echo HATA: "%EXENAME%" bu klasorde bulunamadi.
    echo Lutfen kurulum.bat dosyasini %EXENAME% ile ayni klasorde calistirin.
    echo.
    pause
    exit /b 1
)

echo [1/3] Program dosyalari kopyalaniyor...
if not exist "%APPDIR%" mkdir "%APPDIR%" >nul 2>&1
copy /Y "%SRC%" "%APPDIR%\%EXENAME%" >nul
if errorlevel 1 (
    echo HATA: Dosya kopyalanamadi.
    pause
    exit /b 1
)

echo [2/3] Kisayollar olusturuluyor...
> "%PS1%" echo $ws = New-Object -ComObject WScript.Shell
>> "%PS1%" echo $target = "%APPDIR%\%EXENAME%"
>> "%PS1%" echo $desktop = [Environment]::GetFolderPath('Desktop')
>> "%PS1%" echo $s1 = $ws.CreateShortcut([System.IO.Path]::Combine($desktop, 'Nobet Planlama Sistemi.lnk'))
>> "%PS1%" echo $s1.TargetPath = $target
>> "%PS1%" echo $s1.WorkingDirectory = "%APPDIR%"
>> "%PS1%" echo $s1.IconLocation = $target
>> "%PS1%" echo $s1.Save()
>> "%PS1%" echo $programs = [Environment]::GetFolderPath('Programs')
>> "%PS1%" echo $s2 = $ws.CreateShortcut([System.IO.Path]::Combine($programs, 'Nobet Planlama Sistemi.lnk'))
>> "%PS1%" echo $s2.TargetPath = $target
>> "%PS1%" echo $s2.WorkingDirectory = "%APPDIR%"
>> "%PS1%" echo $s2.IconLocation = $target
>> "%PS1%" echo $s2.Save()

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" >nul 2>&1
del "%PS1%" >nul 2>&1

echo [3/3] Kurulum tamamlandi.
echo.
echo Program su konuma kuruldu: %APPDIR%
echo Masaustunde ve Baslat menusunde "Nobet Planlama Sistemi" kisayolu olusturuldu.
echo.
echo NOT: Ilk acilista Windows SmartScreen bir uyari gosterebilir
echo      ("Windows bilgisayarinizi korudu"). Bu normaldir, imzasiz
echo      bir program oldugu icin cikar. "Daha fazla bilgi" -^> "Yine de calistir"
echo      secerek devam edebilirsiniz.
echo.

choice /M "Programi simdi baslatmak ister misiniz"
if errorlevel 2 goto :sonend
start "" "%APPDIR%\%EXENAME%"

:sonend
echo.
echo Kurulum penceresini kapatabilirsiniz.
pause
endlocal
