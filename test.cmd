@echo off
setlocal
set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" exit /b 1
"%CSC%" /nologo /codepage:65001 /target:exe /out:"%~dp0tests\FormulaEngineTests.exe" "%~dp0FormulaEngine.cs" "%~dp0tests\FormulaEngineTests.cs"
if errorlevel 1 exit /b 1
"%~dp0tests\FormulaEngineTests.exe"
if errorlevel 1 exit /b 1
"%CSC%" /nologo /codepage:65001 /target:exe /out:"%~dp0tests\CsvFileTests.exe" "%~dp0CsvFile.cs" "%~dp0tests\CsvFileTests.cs"
if errorlevel 1 exit /b 1
"%~dp0tests\CsvFileTests.exe"
if errorlevel 1 exit /b 1
"%CSC%" /nologo /codepage:65001 /target:exe /main:DinkCel.OpenFileTests /out:"%~dp0tests\OpenFileTests.exe" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Xml.Linq.dll /reference:System.Core.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll "%~dp0DesktopApp.cs" "%~dp0SpreadsheetFeatures.cs" "%~dp0XlsxFile.cs" "%~dp0XlsxStyles.cs" "%~dp0FormulaEngine.cs" "%~dp0ThemePalette.cs" "%~dp0CsvFile.cs" "%~dp0tests\OpenFileTests.cs"
if errorlevel 1 exit /b 1
"%~dp0tests\OpenFileTests.exe"
if errorlevel 1 exit /b 1
"%CSC%" /nologo /codepage:65001 /target:exe /main:DinkCel.XlsxFileTests /out:"%~dp0tests\XlsxFileTests.exe" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Xml.Linq.dll /reference:System.Core.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll "%~dp0DesktopApp.cs" "%~dp0SpreadsheetFeatures.cs" "%~dp0XlsxFile.cs" "%~dp0XlsxStyles.cs" "%~dp0FormulaEngine.cs" "%~dp0ThemePalette.cs" "%~dp0CsvFile.cs" "%~dp0tests\XlsxFileTests.cs"
if errorlevel 1 exit /b 1
"%~dp0tests\XlsxFileTests.exe"
if errorlevel 1 exit /b 1
"%CSC%" /nologo /codepage:65001 /target:exe /main:DinkCel.MultiSheetTests /out:"%~dp0tests\MultiSheetTests.exe" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Xml.Linq.dll /reference:System.Core.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll "%~dp0DesktopApp.cs" "%~dp0SpreadsheetFeatures.cs" "%~dp0XlsxFile.cs" "%~dp0XlsxStyles.cs" "%~dp0FormulaEngine.cs" "%~dp0ThemePalette.cs" "%~dp0CsvFile.cs" "%~dp0tests\MultiSheetTests.cs"
if errorlevel 1 exit /b 1
"%~dp0tests\MultiSheetTests.exe"
if errorlevel 1 exit /b 1
