# Prompt: Vista de Arbol RBS para Apurisk

Implementa la vista de arbol RBS para el add-in Excel `Apurisk` (Apurisk.XlamShell). Actualmente el boton "Arbol RBS" del Ribbon ejecuta una MsgBox placeholder. Quiero reemplazarlo por un `UserForm` de VBA con un canvas de arbol interactivo.

## Contexto del proyecto

- Add-in tipo `.xlam` con VBA. Codigo en `src/Apurisk.XlamShell/vba/`.
- El Ribbon tiene boton `Arbol RBS` con `onAction="Apurisk_OpenRbsTree"`, que llama a `ApuriskRibbonCallbacks.Apurisk_OpenRbsTree` → `ApuriskBowTieModule.Apurisk_ShowRbsTreePlaceholder`.
- Los datos de RBS viven en la hoja `Apurisk_RBS` con columnas: `CodigoRBS`, `Nombre`, `PadreRBS`, `Nivel`, `Descripcion`. Los codigos usan notacion de punto: `1`, `1.1`, `1.2`, `1.2.1`, etc.
- Los datos de riesgos estan en la tabla maestra del usuario (rangos configurados desde la vista "Ingresar Valores BowTie", guardados en `Apurisk_Config` y mapeados en `Apurisk_RiskMaster_Map`).

## Requisitos de la vista

### 1. Ventana emergente
Un `UserForm` que se abra al hacer clic en "Arbol RBS". Debe ser redimensionable, maximizable y modal (`vbModeless` para que el usuario pueda ver Excel detras si quiere).

### 2. Arbol jerarquico tipo PrecisionTree / arbol de decision
- El nodo raiz (codigo RBS nivel 1, ej. `1`) aparece al lado izquierdo.
- Al hacer clic en un nodo, se despliegan sus subcategorias a la derecha (hijos por `PadreRBS`).
- Cada nivel del arbol avanza horizontalmente (layout izquierda → derecha).
- Los nodos hoja (ultimo nivel del RBS) deben mostrar, conectados a la derecha, los codigos de riesgo asociados con su texto de descripcion (leidos de la tabla maestra del usuario, filtrando por el codigo RBS del nodo).

### 3. Canvas navegable
- Fondo blanco (`Canvas` o `PictureBox` o `Frame` grande con scroll).
- **Zoom/pan** con la ruedita del mouse (`MouseWheel` para zoom in/out centrado en la posicion del cursor; clic sostenido + arrastre para pan).
- Barras de scroll horizontal y vertical (`ScrollBars`) para moverse cuando el arbol no cabe en la ventana.

### 4. Botones flotantes
En esquina superior izquierda del canvas:
- **"Colapsar todo"**: colapsa todos los nodos del arbol.
- **"Expandir todo"**: expande recursivamente todos los nodos.

### 5. Dibujo de nodos y conexiones
- Cada nodo se dibuja como un rectangulo redondeado con el codigo + nombre RBS dentro, usando colores de la paleta de Apurisk (definidos en `ApuriskBowTieIntake`).
- Las lineas de conexion entre nodos deben ser visibles (lineas horizontales y verticales estilo arbol).
- Al hacer clic en un nodo hoja de RBS, los riesgos asociados aparecen anclados a la derecha con su ID y descripcion.

### 6. Integracion
- Reemplazar `Apurisk_ShowRbsTreePlaceholder` en `ApuriskBowTieModule.bas` para que lance el UserForm.
- Crear un nuevo modulo `ApuriskRbsTreeView.bas` con la logica de construccion del arbol, lectura de datos desde `Apurisk_RBS` y la tabla maestra, y renderizado.
- Crear el UserForm `frmRbsTree` con los controles necesarios.

### 7. Manejo de errores
Si no hay hoja `Apurisk_RBS` o si esta vacia, mostrar un mensaje informativo y no abrir la vista.

### 8. Exportabilidad
El UserForm debe poder exportarse como `.frm` para control de versiones junto a los `.bas`.

## Reglas de estilo
Seguir las convenciones existentes: `Option Explicit`, prefijo `Apurisk_` en procedimientos publicos, nombres descriptivos, y paleta de colores del modulo `ApuriskBowTieIntake`.

## Resultado esperado
Un modulo `.bas` nuevo y un `.frm` de UserForm que, al ser importados en el `.xlam`, conviertan el boton "Arbol RBS" en una vista completa, interactiva y navegable del arbol jerarquico de riesgos.
