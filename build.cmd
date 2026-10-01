@echo off
setlocal
set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
  echo Khong tim thay trinh bien dich C# cua .NET Framework.
  exit /b 1
)
"%CSC%" /nologo /codepage:65001 /target:winexe /platform:anycpu /win32icon:"%~dp0DinkCel.ico" /out:"%~dp0DinkCel.exe" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Xml.Linq.dll /reference:System.Core.dll "%~dp0DesktopApp.cs" "%~dp0FormulaEngine.cs" "%~dp0ThemePalette.cs" "%~dp0CsvFile.cs" "%~dp0VersionInfo.cs"
if errorlevel 1 exit /b 1
echo Da tao "%~dp0DinkCel.exe"
