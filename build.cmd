@echo off
setlocal
set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
  echo Khong tim thay trinh bien dich C# cua .NET Framework.
  exit /b 1
)
"%CSC%" /nologo /codepage:65001 /target:winexe /platform:anycpu /win32icon:"%~dp0DinkCel.ico" /out:"%~dp0DinkCel.exe" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Xml.Linq.dll /reference:System.Core.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll /reference:"%~dp0vendor\NPOI.dll" /reference:"%~dp0vendor\PdfSharp.dll" /reference:System.Windows.Forms.DataVisualization.dll /resource:"%~dp0vendor\NPOI.dll",NPOI.dll /resource:"%~dp0vendor\PdfSharp.dll",PdfSharp.dll /resource:"%~dp0vendor\ICSharpCode.SharpZipLib.dll",ICSharpCode.SharpZipLib.dll /resource:"%~dp0vendor\THIRD_PARTY_LICENSES.txt",ThirdPartyLicenses.txt "%~dp0DesktopApp.cs" "%~dp0WorkbookFeatures.cs" "%~dp0EmbeddedDependencies.cs" "%~dp0XlsFile.cs" "%~dp0OdsFile.cs" "%~dp0PdfFile.cs" "%~dp0SpreadsheetV3Ui.cs" "%~dp0SpreadsheetOutputUi.cs" "%~dp0SpreadsheetFeatures.cs" "%~dp0XlsxFile.cs" "%~dp0XlsxStyles.cs" "%~dp0XlsxCharts.cs" "%~dp0FormulaEngine.cs" "%~dp0ThemePalette.cs" "%~dp0CsvFile.cs" "%~dp0VersionInfo.cs"
if errorlevel 1 exit /b 1
echo Da tao "%~dp0DinkCel.exe"
