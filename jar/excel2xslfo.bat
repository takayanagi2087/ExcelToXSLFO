@echo off
set SCRIPT_DIR=%~dp0
if "%JAVA_HOME%" == "" goto noJavaHome
if not exist "%JAVA_HOME%\bin\java.exe" goto noJavaHome
if "%JAVACMD%" == "" set JAVACMD=%JAVA_HOME%\bin\java
goto runApp
:noJavaHome
if "%JAVACMD%" == "" set JAVACMD=java
:runApp
"%JAVACMD%" %JAVAOPTS% -cp %SCRIPT_DIR%excel2xslfo.jar;%SCRIPT_DIR%lib\* exeltoxslfo.ExcelToXSLFO %*
