@echo off
REM Establece la ruta de origen y destino
set origen=%~dp0main.exe
set destino=C:\Registro de novedades y cambios de turnos TK\main.exe

REM Comprobar si el archivo existe en el destino y sobrescribirlo
if exist "%destino%" (
    copy /y "%origen%" "%destino%"
    echo El archivo ha sido sobrescrito con exito.
) else (
    copy "%origen%" "%destino%"
    echo El archivo ha sido copiado con exito.
)

pause
