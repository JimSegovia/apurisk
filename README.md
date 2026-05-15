# Apurisk

Apurisk es una extension COM para Excel orientada a estadistica y gestion de riesgos. El primer modulo es **Analisis BowTie**.

## Requisitos

- Windows con .NET Framework 4.0+
- Excel 2013 / 2016 / Office 365 (con PIA de Office en GAC)
- PowerShell 5+

## Arranque rapido

```powershell
# 1. Compilar
.\scripts\build.ps1 -Configuration Debug

# 2. Registrar en Excel
.\scripts\register-addin.ps1 -Configuration Debug

# 3. Abrir Excel -> pestana "Apurisk" -> boton "Ingresar valores"
```

Si el add-in no aparece:

```powershell
.\scripts\enable-addin.ps1
```

## Despues de hacer cambios en el codigo

Cada vez que modifiques archivos `.cs` en `src/Apurisk.ExcelAddIn/`, ejecuta estos 3 pasos:

```powershell
# 1. Recompilar
.\scripts\build.ps1 -Configuration Debug

# 2. Si Excel esta abierto, CIERRALO antes de continuar

# 3. Re-registrar (el script sobreescribe las entradas del registro)
.\scripts\register-addin.ps1 -Configuration Debug

# 4. Abrir Excel y probar
```

> **Importante**: Siempre cierra Excel antes de recompilar/registrar, o el DLL nuevo no se cargara.

## Flujo del boton principal

El boton **"Ingresar valores"** en la pestana `Apurisk` abre el formulario de captura de rangos.

Dentro del formulario:

1. Haz clic en un campo blanco (se resalta en azul).
2. Presiona **"Seleccionar rango"**.
3. Excel abre el selector nativo de rangos.
4. Presiona **"Aceptar"** para guardar. Los datos se persisten automaticamente.

Campos:

- Nombre RBS
- Codigo RBS
- Seleccion automatica
- ID
- TOP
- Codigo RBS del riesgo
- Nombre RBS del riesgo
- Descripcion del riesgo
- Causas clave
- Impacto / efecto potencial
- Probabilidad
- Impacto
- Gravedad
- Medidas de mitigacion
- Persona responsable

Ademas hay una zona de **Impactos adicionales** con los botones `+ Impacto` y `- Impacto`.

## Arbol RBS

El boton **"Arbol RBS"** abre una ventana en pantalla completa con el arbol jerarquico de categorias RBS, construido a partir de los rangos configurados en "Ingresar valores".

- Los nodos muestran el codigo RBS y su nombre.
- **Click en un nodo**: expande o colapsa sus subcategorias.
- Los nodos hoja (sin subcategorias) muestran los riesgos asociados (ID + descripcion), leidos de los rangos de la tabla maestra.
- Botones **"Expandir todo"** / **"Colapsar todo"** para control global.
- Layout tipo arbol de decision: izquierda → derecha, hijos centrados verticalmente respecto al padre.
- Fondo blanco, paleta de colores sobria tipo Office.

## Estructura del proyecto

```
src/
  Apurisk.Core/          Modelos de dominio (RBS, BowTie, RiskItem)
  Apurisk.Application/   Capa de aplicacion (controllers, gateways)
  Apurisk.ExcelAddIn/    Add-in COM para Excel
    Connect.cs           Punto de entrada COM + ribbon callbacks
    BowTieBootstrapper.cs   Orquestador de acciones
    Excel/               Gateway de comunicacion con Excel
    Forms/               Formularios Windows Forms
      BowTieIntakeForm.cs   Captura de rangos
      RbsTreeForm.cs        Arbol RBS interactivo
    Ribbon/              XML del ribbon
  Apurisk.XlamShell/     Shell VBA (legacy, reemplazado por el add-in C#)
scripts/
  build.ps1              Compila todos los proyectos
  register-addin.ps1     Registra el add-in COM en el registro de Windows
  enable-addin.ps1       Reactiva el add-in si Excel lo deshabilito
docs/
  apurisk_arquitectura_inicial.md
  apurisk_architecture_log.md
```

## Estado actual

- Add-in COM en C# con ribbon nativo.
- Formulario Windows Forms para captura de rangos RBS y tabla maestra.
- Selector visual de rangos usando el InputBox nativo de Excel.
- Persistencia interna via CustomXMLParts (datos invisibles dentro del archivo, sin hojas extra).
- Arbol RBS interactivo con navegacion jerarquica y riesgos asociados en nodos hoja.
