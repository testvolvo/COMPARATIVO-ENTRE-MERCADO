# Comparativo de Modelos — GitHub Pages + SharePoint sincronizado

Esta versión no contiene la base dentro del HTML. `index.html` carga `modelos.json` y verifica `version.json` al abrir.

## Primera configuración

1. En SharePoint, abre la carpeta de la matriz y usa **Agregar acceso directo a OneDrive** o **Sincronizar**.
2. En el Explorador de Windows, comprueba que el Excel tenga el icono verde o selecciona **Mantener siempre en este dispositivo**.
3. Clona o abre este repositorio con GitHub Desktop.
4. Ejecuta `actualizador/Configurar_Actualizador.bat` y selecciona el Excel sincronizado.
5. Ejecuta `actualizador/Actualizar_Comparativo.bat`.
6. El script genera `modelos.json`, actualiza `version.json`, crea un commit y hace push. GitHub Pages se actualizará después del despliegue.

## Uso habitual

Cuando cambie el Excel de SharePoint:

1. Espera a que OneDrive termine de sincronizar.
2. Ejecuta `actualizador/Actualizar_Comparativo.bat`.
3. Los demás usuarios no necesitan OneDrive: solo abren la página de GitHub Pages.

El comparador carga la última base guardada en el navegador y consulta primero `version.json`. Solo vuelve a descargar `modelos.json` si cambió su versión.

## Seguridad y columnas publicadas

GitHub Pages y los archivos JSON pueden ser visibles públicamente. El actualizador **no exporta todas las columnas del Excel**: solo las enumeradas en `columnas_publicas.json`.

Antes del primer push, revisa especialmente los campos de precios y equivalencias. Para dejar de publicar una columna, elimínala de `columnas_publicas.json` y también del comparador si aparece en la matriz.

## Traspaso a otra persona

No hay rutas personales guardadas en el repositorio. La configuración queda en:

`%LOCALAPPDATA%\ComparativoModelosVolvo\configuracion.json`

La nueva persona solo debe:

1. Tener acceso a la carpeta de SharePoint.
2. Sincronizarla con OneDrive.
3. Clonar el mismo repositorio con GitHub Desktop.
4. Ejecutar `Configurar_Actualizador.bat` y seleccionar su propia copia sincronizada.
5. Ejecutar `Actualizar_Comparativo.bat`.

No necesita cambiar el HTML ni los scripts. Debe tener permiso de escritura en el repositorio de GitHub.

## Si Git no publica automáticamente

El script igualmente dejará actualizados `modelos.json` y `version.json`. Abre GitHub Desktop, escribe un resumen, pulsa **Commit to main** y luego **Push origin**.

## Requisitos

- Windows con Microsoft Excel instalado.
- OneDrive corporativo iniciado y la carpeta sincronizada.
- Git o GitHub Desktop para publicar.
- El Excel debe conservar la hoja `BBDD MARCAS` y sus encabezados.
