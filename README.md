# Apurisk

Apurisk es una base de extension para Excel orientada a estadistica y gestion de riesgos. El primer modulo sera **Analisis BowTie**.

## Arranque rapido

Base COM experimental:

```powershell
.\scripts\build.ps1
```

Shell XLAM actual:

```powershell
src\Apurisk.XlamShell\
```

## Estado actual

La base incluye:

- Shell `XLAM + VBA` con Ribbon `Apurisk`.
- Grupo inicial `Analisis BowTie`.
- Boton principal `Analisis BowTie` para iniciar la captura del modulo.
- Botones `Arbol RBS`, `Abrir BowTie`, `Validar` e `Insertar valores`.
- Modelo inicial de RBS, tabla maestra y BowTie.
- Popup unificado para capturar rangos de RBS y tabla maestra, con estilo mas cercano al look nativo de Excel/VBA.
- Seleccion visual de rangos usando el selector nativo de Excel.
- Persistencia de rangos y del ID elegido para poder corregir valores luego.
- `Office RibbonX Editor` instalado localmente para editar `customUI`.

## Herramienta RibbonX

Usaremos `Office RibbonX Editor` de `fernandreu` como editor oficial de la parte `customUI` del add-in:

- Repositorio oficial: [fernandreu/office-ribbonx-editor](https://github.com/fernandreu/office-ribbonx-editor)
- Release usada para esta base: `v1.9.0`
- Variante instalada: `OfficeRibbonXEditor-NETFramework-Binaries.zip`
- Ruta local: [tools/OfficeRibbonXEditor](D:/JimRisk/RiskCode/Apurisk/tools/OfficeRibbonXEditor)
- Ejecutable: [OfficeRibbonXEditor.exe](D:/JimRisk/RiskCode/Apurisk/tools/OfficeRibbonXEditor/OfficeRibbonXEditor.exe)
- Script de apertura: [scripts/open-ribbonx-editor.ps1](D:/JimRisk/RiskCode/Apurisk/scripts/open-ribbonx-editor.ps1)

Detalles importantes tomados del repositorio oficial:

- El editor es una herramienta standalone para editar la parte `Custom UI` de archivos Office abiertos como `xlsx`, `xlsm`, `xlam`, `pptm` o `docx`.
- La release `v1.9.0` recomienda usar el instalador o binarios `.NET Framework` si hay duda.
- Desde `v2.0` ya no se soporta `.NET Framework`; la ultima version `.NET Framework` indicada por el proyecto es `v1.9`.
- El archivo Office se trata como un `.zip`; por eso conviene cerrar Excel antes de editar `customUI` y luego volver a abrir el archivo para ver cambios limpios.

## Flujo XLAM

Para ver `Apurisk` en Excel con el enfoque actual:

1. Crear un libro nuevo en Excel.
2. Guardarlo como `Apurisk.xlam`.
3. Importar los modulos `.bas` desde [src/Apurisk.XlamShell/vba](D:/JimRisk/RiskCode/Apurisk/src/Apurisk.XlamShell/vba).
4. Importar la UserForm [frmApuriskBowTieIntake.frm](D:/JimRisk/RiskCode/Apurisk/src/Apurisk.XlamShell/forms/frmApuriskBowTieIntake.frm).
5. Verificar que [frmApuriskBowTieIntake.frx](D:/JimRisk/RiskCode/Apurisk/src/Apurisk.XlamShell/forms/frmApuriskBowTieIntake.frx) este en la misma carpeta durante la importacion.
6. Abrir `Apurisk.xlam` con [OfficeRibbonXEditor.exe](D:/JimRisk/RiskCode/Apurisk/tools/OfficeRibbonXEditor/OfficeRibbonXEditor.exe).
7. Insertar o reemplazar el XML con [customUI14.xml](D:/JimRisk/RiskCode/Apurisk/src/Apurisk.XlamShell/customUI/customUI14.xml).
8. Guardar el `xlam`, cerrar Excel si estaba abierto y volver a abrirlo.
9. Activar el complemento `Apurisk.xlam` en Excel.

## Flujo actual del boton principal

El boton `Analisis BowTie` abre un popup unificado llamado `Analisis BowTie - Ingresar Valores`.

Dentro del popup:

1. Seleccionas un cuadro blanco.
2. Presionas `cargar rangos de celda`.
3. Excel abre el selector visual de rangos.
4. El rango queda guardado y se vuelve a mostrar cuando reabres el popup.

Campos principales:

- `Nombre RBS`
- `Codigo RBS`
- `Seleccion automatica`
- `ID`
- `TOP`
- `Codigo RBS del riesgo`
- `Nombre RBS del riesgo`
- `Descripcion del riesgo`
- `Causas clave`
- `Impacto / efecto potencial`
- `Probabilidad`
- `Impacto`
- `Gravedad`
- `Medidas de mitigacion`
- `Persona responsable`
- `ID del riesgo a analizar`

Ademas hay una zona de `Impactos adicionales` con `agregar impacto`.

La arquitectura inicial esta documentada en [docs/apurisk_arquitectura_inicial.md](/D:/JimRisk/RiskCode/Apurisk/docs/apurisk_arquitectura_inicial.md) y el log continuo en [docs/apurisk_architecture_log.md](/D:/JimRisk/RiskCode/Apurisk/docs/apurisk_architecture_log.md).
