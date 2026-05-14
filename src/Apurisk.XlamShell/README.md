# Apurisk XLAM Shell

Esta carpeta contiene la shell inicial de `Apurisk` para Excel usando `XLAM + VBA`.

## Objetivo

La shell XLAM nos permite avanzar rapido en:

- pestana `Apurisk` en el Ribbon;
- comandos del modulo `Analisis BowTie`;
- creacion de hojas base;
- popup unificado de captura con look mas nativo de Excel/VBA mientras definimos la vista BowTie completa.

## Estructura

```text
Apurisk.XlamShell/
  customUI/
    customUI14.xml
  forms/
    frmApuriskBowTieIntake.frm
    frmApuriskBowTieIntake.frx
  vba/
    ApuriskBootstrap.bas
    ApuriskRibbonCallbacks.bas
    ApuriskWorkbookSetup.bas
    ApuriskBowTieModule.bas
    ApuriskState.bas
    ApuriskBowTieIntake.bas
```

## Como montarlo en Excel

1. Crear un libro nuevo en Excel.
2. Guardarlo como `Apurisk.xlam`.
3. Importar los modulos `.bas` desde `vba/`.
4. Importar `forms/frmApuriskBowTieIntake.frm`.
5. Verificar que `forms/frmApuriskBowTieIntake.frx` este en la misma carpeta para que Excel lo tome durante la importacion.
6. Abrir `Apurisk.xlam` con `Office RibbonX Editor`.
7. Insertar `customUI/customUI14.xml` usando el editor.
8. Guardar el archivo y cerrarlo.
9. Cerrar y volver a abrir Excel.
10. Activar el complemento `Apurisk.xlam`.

## Herramienta usada

Usaremos `Office RibbonX Editor` como editor de `customUI`.

- Repositorio: [fernandreu/office-ribbonx-editor](https://github.com/fernandreu/office-ribbonx-editor)
- Version instalada: `v1.9.0`
- Ruta local: [OfficeRibbonXEditor.exe](D:/JimRisk/RiskCode/Apurisk/tools/OfficeRibbonXEditor/OfficeRibbonXEditor.exe)
- Script de apertura: [open-ribbonx-editor.ps1](D:/JimRisk/RiskCode/Apurisk/scripts/open-ribbonx-editor.ps1)

Notas practicas:

- El proyecto recomienda el binario `.NET Framework` si hay duda.
- Conviene cerrar Excel antes de editar `customUI`, guardar en el editor y luego volver a abrir el add-in.
- El editor trabaja sobre el paquete Office como si fuera un `.zip`, asi que es mejor evitar tener el mismo archivo abierto en Excel mientras se cambia el Ribbon.

## Alcance de esta fase

Esta shell no reemplaza la arquitectura por capas. Solo reemplaza el punto de entrada para que podamos avanzar sin depender del arranque COM.

La logica de negocio futura debe seguir separada de la UI del Ribbon y de las macros de Excel.

## Flujo actual

El boton principal del Ribbon ahora es `Analisis BowTie`.

Al presionarlo, se abre un popup unificado llamado `Analisis BowTie - Ingresar Valores`.

La interaccion esperada es:

1. Hacer clic en un cuadro del popup.
2. Presionar `cargar rangos de celda`.
3. Seleccionar el rango visualmente en Excel.
4. Aceptar para guardar.

Los rangos seleccionados quedan guardados en `Apurisk_Config` y vuelven a aparecer si el usuario reabre el popup para corregir algo.

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

Campos dinamicos:

- `Cat. Impacto 1`
- `Cat. Impacto 2`
- y asi sucesivamente con el boton `agregar impacto`
