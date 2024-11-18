El Registro de Novedades y Cambios de Turnos es una aplicación sencilla creada para registrar las novedades en un archivo excel por medio de un formulario que permita validar datos y evitar que por accidente puedan ser eliminados. Al estar desarrollado en Python es mucho mas veloz que implementar una Macro en excel

No requiere ninguna instalación si se ejecuta el main.exe dentro de la carpeta dist.
En caso de querer revisar el código y empaquetarlo necesitaran tener todas las librerías instaladas y empaquetarlo con pyinstaller pueden hacerlo con el siguiente comando:

```
pyinstaller --onefile --windowed main.py
```

El ejecutable puede tener un archivo de configuración en la misma carpeta con el nombre "path_base" sin ninguna extension donde se puede colocar la ruta del excel con el nombre.
Por ejemplo: 

```
C:\Users\user\workspace\registroNovedadesTk\assets\PLANILLA NOVEDADES PERSONAL ABORDO.xlsx
```

En caso de que este archivo no exista al ejecutarse la aplicación creara un por default en la misma carpeta, con las hojas y las tablas necesaria pero vacías

Recomiendo descargar el archivo main.exe y agregarle el archivo de configuración "path_base" y colocarlo en la unidad D: en caso de tener la pc y crear un acceso directo en el escritorio.