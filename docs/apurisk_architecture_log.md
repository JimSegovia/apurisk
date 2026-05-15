# Apurisk - log de arquitectura

## 2026-05-14 10:05 -05:00

Se pivoto el punto de entrada de Excel desde una prueba `COM/.NET` artesanal hacia una shell `XLAM + VBA`.

### Motivo

Excel detectaba el complemento COM, pero abortaba la carga antes de ejecutar `OnConnection`. Eso bloqueaba el arranque del producto y nos frenaba en una capa de infraestructura en vez de avanzar con el modulo BowTie.

### Decision

- Mantener la arquitectura por capas como direccion del producto.
- Usar `XLAM` como shell inicial de Excel.
- Conservar `Apurisk.Core` y `Apurisk.Application` como referencia de dominio y orquestacion futura.
- Priorizar ahora flujo, UX y estructura del modulo `Analisis BowTie`.

### Se trabajo en

- crear `src/Apurisk.XlamShell/`;
- definir `customUI14.xml` para la pestana `Apurisk`;
- crear modulos VBA base para bootstrap, Ribbon y hojas de trabajo;
- documentar el montaje manual del `XLAM`.

### Se trabajara ahora

- ventana de configuracion de tabla maestra;
- captura de columnas `ID`, `RBS`, `Nombre`, `Descripcion`;
- lectura del catalogo RBS;
- primera navegacion tipo arbol por RBS y riesgos asociados.

## 2026-05-14 10:36 -05:00

Se adopto `Office RibbonX Editor` como herramienta oficial para editar `customUI` del `Apurisk.xlam`.

### Motivo

Necesitamos un flujo estable y repetible para insertar y mantener el Ribbon del add-in sin depender de herramientas ambiguas o pasos manuales poco documentados.

### Decision

- usar `Office RibbonX Editor` del repositorio `fernandreu/office-ribbonx-editor`;
- instalar la release `v1.9.0` variante `.NET Framework Binaries`;
- dejar la herramienta versionada a nivel de workspace en `tools/OfficeRibbonXEditor/`;
- documentar el flujo en `README.md` y `src/Apurisk.XlamShell/README.md`.

### Se trabajo en

- descargar e instalar `OfficeRibbonXEditor-NETFramework-Binaries.zip`;
- crear `scripts/open-ribbonx-editor.ps1`;
- actualizar la documentacion del flujo `XLAM + RibbonX`.

### Se trabajara ahora

- ensamblar el primer `Apurisk.xlam`;
- importar modulos VBA;
- insertar `customUI14.xml`;
- validar que la pestana `Apurisk` aparezca en Excel.

## 2026-05-14 10:50 -05:00

Se convirtio el boton principal del Ribbon en el arranque del modulo `Analisis BowTie`.

### Motivo

El flujo ya no debia quedarse en placeholders de base. Necesitabamos una captura inicial real para RBS, tabla maestra y seleccion del riesgo a analizar.

### Decision

- reemplazar `Crear base` por `Analisis BowTie`;
- usar el titulo de flujo `Analisis BowTie - Ingresar Valores`;
- dividir la captura en `Carga de RBS` y `Carga de Tabla Maestra de Riesgos`;
- pedir `Codigo RBS del riesgo` como campo obligatorio y `Nombre RBS del riesgo` como opcional.

### Se trabajo en

- actualizar `customUI14.xml` y callbacks del Ribbon;
- crear `ApuriskBowTieIntake.bas`;
- validar columnas obligatorias antes de continuar;
- guardar configuracion y mapeos en el libro activo del usuario.

### Se trabajara ahora

- armar la vista tipo arbol usando el `Codigo RBS`;
- filtrar riesgos por jerarquia RBS;
- abrir el riesgo elegido en la futura vista BowTie.

## 2026-05-14 12:32 -05:00

Se deshizo el enfoque de hoja `Apurisk_BowTie_Intake` y se reemplazo por un popup `UserForm` real.

### Motivo

La vista en hoja no cumplia con la expectativa del modulo: el usuario queria un popup tipo formulario, con cajas blancas clicables y el selector de rangos ocurriendo desde esa misma ventana.

### Decision

- eliminar la idea de una hoja dedicada para la captura;
- usar un `UserForm` importable como `frmApuriskBowTieIntake.frm/.frx`;
- mantener la seleccion visual de rangos con `Application.InputBox Type:=8`;
- mantener persistencia en `Apurisk_Config` para que los valores se conserven al reabrir el popup.

### Se trabajo en

- revertir referencias a `Apurisk_BowTie_Intake`;
- reescribir `ApuriskBowTieIntake.bas` como soporte del popup;
- generar `frmApuriskBowTieIntake.frm` y `frmApuriskBowTieIntake.frx`;
- actualizar documentacion y flujo de importacion.

### Se trabajara ahora

- validar en Excel el comportamiento real del popup;
- afinar posiciones y tamano de controles segun la experiencia visual;
- mejorar la edicion de impactos multiples si hace falta.

## 2026-05-14 12:46 -05:00

Se mejoro el estilo visual del popup de captura para acercarlo al look nativo de Excel/VBA.

### Motivo

La primera version del popup ya funcionaba, pero se seguia viendo demasiado generada. La meta de esta pasada fue hacer que el formulario se sintiera mas profesional y mas propio del ecosistema de Excel.

### Decision

- usar un fondo de dialogo gris tipo Office/VBA;
- agrupar campos con `Frame` nativos;
- mantener botones estandar sin colores artificiales;
- conservar la logica de seleccion de rangos y persistencia sin cambiar el flujo.

### Se trabajo en

- ajustar `scripts/generate-bowtie-userform.ps1`;
- regenerar `frmApuriskBowTieIntake.frm/.frx`;
- mantener el popup como artefacto importable para el `xlam`.

### Se trabajara ahora

- validar visualmente el popup dentro de Excel;
- hacer ajustes finos de espaciado o tamano si los ves necesarios;
- seguir con el arbol RBS una vez que la base visual quede comoda.

## 2026-05-15 - Migracion a C# WinForms y CustomDocumentProperties

Se migro el add-in de VBA UserForm a C# WinForms y se cambio la persistencia.

### Persistencia: CustomDocumentProperties

Los datos de configuracion (rangos, impactos) se almacenan como **propiedades personalizadas del documento** (`Workbook.CustomDocumentProperties`):

- **API**: `CustomDocumentProperties.Add(Name, LinkToContent, Type, Value)`
- **Tipo**: `msoPropertyTypeString` (4)
- **Formato del nombre**: `Apur_Field_X` (ej: `Apur_Field_RbsCodeRange`)
- **Lectura**: iterar `foreach (prop in workbook.CustomDocumentProperties)`, buscar por `prop.Name`
- **Escritura**: eliminar propiedad existente (`prop.Delete()`), crear nueva (`Add()`)
- **Límite**: 255 caracteres por valor (suficiente para direcciones de rango)
- **Ventaja**: nativo de Excel, invisible al usuario, sin hojas, persiste con el archivo

### Flujo de guardado

```
BtnAceptar_Click → SaveAllConfig(form)
  ├── foreach field: SaveConfigProp(workbook, "Field." + key, value)
  │     ├── Buscar y eliminar prop existente (Delete)
  │     └── Si value no vacio: Add(propName, false, 4, value)
  └── ImpactFieldCount tambien se guarda
```

### Flujo de carga

```
LoadSavedValues() → foreach key: ReadConfigValue("Field." + key)
  └── Iterar CustomDocumentProperties buscando "Apur_Field_" + key
```

### Intentos previos descartados

1. **Hoja `Apurisk_Config`**: Funcionaba pero el usuario no queria hojas visibles
2. **CustomXMLParts con namespace XML**: Namespace + XPath fallaba con COM late-bound
3. **CustomXMLParts sin namespace XML**: `LoadXML` multiple en mismo part fallaba (solo primera escritura)
4. **CustomDocumentProperties**: ✅ Funciona, nativo, confiable
