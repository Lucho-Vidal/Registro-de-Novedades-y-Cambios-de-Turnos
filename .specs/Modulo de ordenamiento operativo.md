
# Spec / Prompt de Desarrollo: Módulo de Ordenamientos Operativos (RENO)

## 1. Resumen y Objetivo General

Agregar una pestaña adicional denominada **"Ordenamientos"** dentro del módulo de  **Novedades y Cambios** . Esta funcionalidad permitirá gestionar órdenes de servicio diarias para el personal de conducción, permitiendo la selección dinámica de personal, itinerarios, turnos y tramos específicos (incluyendo estaciones intermedias gracias a una nueva importación de datos), control de tiempos de edición por roles, borrado lógico con papelera de recuperación, y generación de un reporte PDF en formato oficial por duplicado listo para imprimir, cortar y firmar.

## 2. Requerimientos de Base de Datos y Estructura (SQLite / Supabase)

### A. Modificación de la Importación de Itinerarios (Soporte de Estaciones Intermedias)

Dado que los turnos actuales solo contemplan las cabeceras (origen/destino principales del bloque), es necesario actualizar el script de importación y el esquema de la tabla de tramos (`turnos_vueltas` o equivalente) para registrar **todas las estaciones intermedias y horarios de paso** de cada tren.

**Esquema de la tabla de tramos actualizado:**

* `id` (PK, INTEGER / UUID)
* `turno_id` / `id_turno` (FK)
* `nro_vuelta` (INTEGER)
* `ref` (VARCHAR, ej:`'PAS.'`,`NULL`)
* `nro_tren` (VARCHAR / INTEGER)
* `origen` (VARCHAR, estación de partida del tramo)
* `hora_salida` (TIME)
* `destino` (VARCHAR, estación de llegada del tramo)
* `hora_llegada` (TIME)
* `observaciones_tramos` (TEXT)
* `sync_status`,`created_at`,`updated_at`

### B. Nueva Tabla: `ordenamientos`

Almacenará el registro histórico de cada orden de servicio emitida.

* `id` (PK, UUID / SERIAL)
* `nro_orden` (VARCHAR, ej:`'PC 0000-0098'` o autoincremental por depósito)
* `fecha_emision` (DATE)
* `hora_emision` (TIME)
* `legajo_personal` (VARCHAR / INT)
* `nombre_personal` (VARCHAR)
* `deposito` (VARCHAR, ej:`'DE AUXILIARES P. Constitución'`)
* `itinerario` (VARCHAR, ej:`'Itinerario H5'`)
* `ciclo_o_turno` (VARCHAR, ej:`'401,1'`)
* `regimen` (ENUM:`'HÁBIL'`,`'SÁBADO'`,`'DOMINGO'`)
* `motivo_orden` (TEXT, ej:`'POR ORDEN DE ELEMENTOS Y RAZONES OPERATIVAS UD. CORRE LO SIGUIENTE:'`)
* `observacion_pie` (TEXT, ej:`'PIDE ORDEN - Aplica artículo 7 del RITO'`)
* `ref_pct` (VARCHAR)
* `emisor` (VARCHAR, usuario logueado)
* `estado` (ENUM:`'activo'`,`'eliminado'`) [Soporte para borrado lógico]
* `created_at`,`updated_at`,`deleted_at`

## 3. Requerimientos de Interfaz de Usuario (UI/UX)

### A. Vista Principal (Pestaña "Ordenamientos")

* **Listado en Tabla:**
  * Mostrar todos los ordenamientos registrados ordenados cronológicamente del **más antiguo primero** (o con opción de alternar orden, pero por defecto histórico ascendente).
  * Columnas sugeridas: N° Orden, Fecha/Hora, Legajo, Conductor, Turno/Itinerario, Emisor, Estado y Acciones (Editar, Ver PDF, Eliminar).
* **Filtros rápidos:** Por fecha, depósito, legajo o número de orden.
* **Acción Principal:** Botón superior *"Nuevo Ordenamiento"* que abre un modal de carga.

### B. Modal / Formulario de Nuevo Ordenamiento

El formulario debe estructurarse de manera intuitiva y conectada con las bases de datos de personal e itinerarios:

1. **Datos Generales (Auto-completados / Editables):**
   * Depósito emisor, Fecha, Hora actual, N° de Orden consecutivo.
   * **Selección de Personal:** Campo de búsqueda por Legajo o Apellido/Nombre que traiga los datos del agente.
   * **Selección de Itinerario y Régimen:** Combobox para elegir el itinerario vigente (ej.`Dic25 / Itinerario H5`) y el régimen operativo (`HÁBIL`,`SÁBADO`,`DOMINGO`).
   * **Turno Base / Diagrama:** Selector del turno asignado (ej.`401,1`).
2. **Carga Dinámica de Tramos y Estaciones Intermedias:**
   * El usuario ingresa o selecciona el **Número de Tren** (ej.`5303` o `1313`).
   * Al ingresar el tren, el sistema consulta la tabla de tramos detallados y despliega un **Combobox de Estaciones Disponibles** por las que pasa dicho tren.
   * **Estación de Origen / Partida:** Al seleccionar la estación de origen de la lista, el sistema autocompleta automáticamente al lado la **hora de salida** correspondiente.
   * **Estación de Destino / Llegada:** Se selecciona la estación de destino (puede ser una estación intermedia o la terminal del tren) y el sistema autocompleta la  **hora de llegada** .
   * Posibilidad de agregar múltiples renglones de tramos (simulando la grilla de la Orden de Servicio de la imagen de referencia^^).
3. **Textos Legales / Observaciones de Pie:**
   * Campos editables para justificación operativa (ej: *"PIDE ORDEN - Aplica artículo 7 del RITO"* ).
   * Campos obligatorios de control: `REF de PCT` y `Emisor`^^.

## 4. Permisos, Roles y Control de Tiempos de Edición

* **Control de Roles:** Solo usuarios con roles autorizados (Controladores de Tránsito, Jefes de Operaciones o Encargados de Personal) podrán crear, editar o eliminar ordenamientos. El campo `Emisor` se registra automáticamente con el usuario actual de la sesión.
* **Tiempo de Edición por Configuración:** La edición o modificación de un ordenamiento emitido debe regirse por la regla global de configuración de RENO (por ejemplo, ventana de tiempo permitida tras su creación, ej. 2 horas). Pasado ese lapso, el registro se bloquea contra ediciones directas y solo permite anulaciones auditadas.
* **Borrado Lógico y Recuperación:**
  * Al eliminar un ordenamiento, el registro no se borra físicamente de la base de datos (`DELETE`), sino que su estado pasa a `eliminado` (`deleted_at = NOW()`).
  * Debe existir un panel de administración o papelera para visualizar registros eliminados y permitir su **restauración** en un click.

## 5. Exportación a PDF e Impresión (Formato Duplicado Oficial)

Basado en la Orden de Servicio oficial adjunta^^, el archivo PDF generado debe cumplir estrictamente con los siguientes parámetros de diseño e impresión:

* **Estructura de Hoja (Duplicado):**
  * Una única hoja tamaño A4 / Oficio dividida en **dos mitades idénticas** (sección superior y sección inferior separadas por una línea de corte punteada) para que el agente firme una mitad que queda en poder de la empresa y conserve el talón correspondiente si aplica.
* **Diseño Gráfico y Cabecera:**
  * Logo de *Trenes Argentinos Operadora Ferroviaria* a la izquierda^^.
  * Título destacado: **ORDEN DE SERVICIO**^^.
  * Cuadro lateral derecho con el código de depósito y número de orden (ej: `PC 0000-0098`), Hora de emisión y Fecha^^.
* **Cuerpo de Datos:**
  * Bloque superior con: Depósito, Legajo, Apellido y Nombre, Itinerario y Mes^^.
  * Texto central de motivo: *"POR ORDEN DE ELEMENTOS Y RAZONES OPERATIVAS UD. CORRE LO SIGUIENTE:"* y código de turno^^.
  * Grilla tabular con los tramos ordenados: Columna de Referencia (`PAS.`), N° de Tren, Origen, Hora Salida, Destino, Hora Llegada y Observaciones^^.
* **Pie de Orden y Firmas:**
  * Línea de notificación: *"PIDE ORDEN"* y mención normativa ( *"Aplica artículo 7 del RITO"* )^^.
  * Casilleros inferiores para `REF de PCT *`, `Emisor *`, y las dos líneas de firma obligatorias: **"Firma del personal"** y **"Firma de Emisor"**^^.
* **Canales de Salida:**
  * Botón **"Imprimir / Descargar PDF"** (genera el documento listo para impresora térmica o láser).
  * Botón **"Enviar por Mail"** (envía el PDF adjunto al correo electrónico registrado del agente o del depósito correspondiente).
