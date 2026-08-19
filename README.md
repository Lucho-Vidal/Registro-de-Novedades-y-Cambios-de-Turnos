El Registro de Novedades y Cambios de Turnos es una aplicación sencilla creada para registrar las novedades en un archivo excel por medio de un formulario que permita validar datos y evitar que por accidente puedan ser eliminados. Al estar desarrollado en Python es mucho mas veloz que implementar una Macro en excel

No requiere ninguna instalación si se ejecuta el main.exe dentro de la carpeta dist.
En caso de querer revisar el código y empaquetarlo necesitaran tener todas las librerías instaladas y empaquetarlo con pyinstaller pueden hacerlo con el siguiente comando:

```
pyinstaller --onefile --windowed main.py
python -m PyInstaller --clean --noconfirm --onefile --windowed --name "RENO" --collect-all bcrypt main.py
```

Antes de empaquetar, instalar las dependencias en el mismo Python que ejecutará PyInstaller. Para incluir correctamente el módulo nativo de bcrypt puede utilizar:

```powershell
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --clean --noconfirm --onefile --windowed --collect-all bcrypt main.py
```

Si se utiliza la especificación del proyecto, ejecutar `python -m PyInstaller main.spec` después de instalar las dependencias.

El ejecutable puede tener un archivo de configuración en la misma carpeta con el nombre "path_base" sin ninguna extensión. Puede contener la ruta de un Excel inicial o directamente la ruta de la base SQLite compartida.
Por ejemplo:

```
C:\Users\user\workspace\registroNovedadesTk\assets\PLANILLA NOVEDADES PERSONAL ABORDO.xlsx
```

En caso de que este archivo no exista al ejecutarse la aplicación creara un por default en la misma carpeta, con las hojas y las tablas necesaria pero vacías

Recomiendo descargar el archivo main.exe y agregarle el archivo de configuración "path_base" y colocarlo en la unidad D: en caso de tener la pc y crear un acceso directo en el escritorio.

## Persistencia SQLite

La aplicación utiliza como fuente principal una base `registro.sqlite` ubicada junto al archivo indicado en `path_base`. También se puede configurar `path_base` directamente con la ruta de una base `.sqlite`. En la primera ejecución, si existe el Excel asociado, importa únicamente las hojas `BASE`, `TipoNovedad`, `NOVEDADES` y `Cambio de Turnos`.

La base puede estar en una carpeta compartida de Windows. Las escrituras utilizan transacciones breves, bloqueo externo y modo journal tradicional; no se utiliza WAL porque las estaciones acceden desde máquinas diferentes. Las copias Excel se generan desde el menú `Archivo > Exportar a Excel`.

Instalación de dependencias:

```text
pip install -r requirements.txt
```
