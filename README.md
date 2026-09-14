# Nodexa · Toma de Inventario

Toma de inventario por **serial** en Excel, pensada para trabajar con **pistola lectora de códigos de barras**: se escanea un serial tras otro sin tocar el teclado ni el mouse, y el archivo compara lo que hay físicamente contra lo que dice el sistema.

Nace de un problema concreto: contar una bodega con cientos de equipos serializados, dos veces al mes, sin planillas paralelas ni fórmulas que se rompen.

> El repositorio incluye una **demo con productos ficticios**. No contiene datos de ninguna empresa.

## Qué hace

- **Escaneo continuo.** Una ventana con el cursor siempre listo: se escanea, se registra y queda esperando el siguiente. Sin marcar celdas ni escribir cantidades.
- **Respuesta inmediata por color y sonido.** Verde si el equipo estaba en el sistema; naranja, rojo o azul según el problema.
- **Ocho resultados distintos**, no solo "está / no está":

  | Resultado | Significado |
  |---|---|
  | `ENCONTRADO` | El serial estaba en el sistema en la sucursal y apareció |
  | `DUPLICADO` | Ya se había escaneado en esta toma; no se cuenta dos veces |
  | `FALTANTE` | Está en el sistema y nunca se escaneó |
  | `NO EXISTE` | Se escaneó algo que el sistema no tiene en ninguna sucursal |
  | `OTRA SUCURSAL` | Existe, pero el sistema lo registra en otra sucursal (indica cuál) |
  | `EN SALIDA` | Tiene Nota de Venta o Picking: no forma parte del stock a contar |
  | `CÓDIGO SKU` | Se escaneó el código del producto, no el serial: se rechaza sin registrar |
  | `AJUSTADO` | Faltante o sobrante justificado a mano, con su motivo |

- **Toma acumulativa.** Se puede cerrar y continuar otro día: el avance queda guardado.
- **Deshacer** el último escaneo, e ingreso manual cuando la etiqueta está dañada.
- **Historial completo** de cada toma cerrada, serial por serial, para responder después "¿este equipo se contó la vez pasada?".
- **Reporte exportable** a un `.xlsx` sin macros, para revisar o enviar desde cualquier PC.
- **Power Query** lee el reporte del sistema (un `.xlsx` que se deja en una carpeta), valida que estén las columnas necesarias y trata el serial siempre como texto, para no perder los ceros iniciales.

## Estructura

```
src/inventario/       código VBA + consulta M del archivo de toma
src/prueba-equipo/    código de la herramienta que valida si un PC puede usarlo
tools/generar_demo.py generador del reporte de stock ficticio
demo/                 paquete listo para probar (se genera, ver más abajo)
```

Los archivos `.bas`, `.txt` (módulos y formularios) y `.m` son el código fuente de verdad; el `.xlsm` se **arma desde ellos** con un script de PowerShell, así el código queda versionable y revisable en Git en vez de encerrado en un binario.

## Probar la demo

Requisitos: Windows, **Excel de escritorio**, Python con `openpyxl` y una pistola lectora (o el teclado, escribiendo el serial y presionando Enter).

```powershell
# 1. Generar el reporte de stock ficticio (570 líneas, 2 sucursales)
python tools\generar_demo.py

# 2. Habilitar temporalmente en Excel:
#    Opciones > Centro de confianza > Configuración de macros >
#    "Confiar en el acceso al modelo de objetos de proyectos de VBA"

# 3. Construir los archivos
powershell -File src\inventario\construir.ps1
powershell -File src\prueba-equipo\construir.ps1

# 4. Desactivar de nuevo esa opción de Excel
```

Queda `demo\Toma_Inventario_Demo.xlsm`. Ábralo, presione **ACTUALIZAR AURORA** y después **ESCANEAR**. Los seriales para probar están en `demo\Aurora\reporte_demo.xlsx`, columna A: algunos son de `Casa Matriz` (se cuentan), otros de `Sucursal Norte` (avisan que son de otra sucursal) y unos pocos tienen Nota de Venta (quedan como `EN SALIDA`).

`demo\Manual de uso.html` es el manual para quien hace el inventario.

## Adaptarlo a otro sistema de stock

El archivo no depende de un sistema en particular: solo necesita un `.xlsx` con una fila por serial. Las columnas que exige están declaradas en `src/inventario/Aurora.m`:

```
Serial · Código · Descripción · Condición · Estado · Bodega · Sucursal · Nota de Venta · Picking
```

Si el reporte trae otros nombres, se ajustan ahí. Si falta una columna, el archivo lo dice con un mensaje claro en vez de fallar en silencio.

La sucursal a contar y las bodegas que entran o no en el conteo se configuran en la hoja **CONFIG** del propio archivo, sin tocar código.

## Decisiones de diseño

- **El código vive fuera del `.xlsm`.** Un binario no se puede revisar en un diff.
- **El registro de escaneos no se borra nunca.** Deshacer marca la fila como anulada; el resultado se recalcula a partir del registro, así que siempre se puede auditar qué pasó.
- **El estado se recalcula contra el reporte actual**, no contra una foto congelada: se puede volver a cargar el reporte a mitad de una toma sin perder lo escaneado.
- **El archivo se guarda con un aviso rojo de "macros bloqueadas"** y solo pasa a verde cuando una macro se ejecuta al abrirlo. Si las macros están bloqueadas en ese PC, el usuario ve el aviso y las instrucciones, en vez de creer que está funcionando.
- **La ruta de OneDrive se resuelve a ruta local.** Excel entrega `https://...` como ubicación del archivo cuando está sincronizado, y Power Query no puede leer desde ahí.
- **Confirmar una toma pide "No" por defecto**, para que un Enter de la pistola no cierre el inventario por accidente.

## Estado

Funcionando en producción en una bodega, con Excel 2024 de 64 bits sobre Windows 11.

Limitaciones conocidas: un equipo a la vez (el archivo no está pensado para edición simultánea), las hojas no están protegidas contra escaneos hechos fuera de la ventana de escaneo, y no avisa si el reporte de stock cargado está desactualizado.
