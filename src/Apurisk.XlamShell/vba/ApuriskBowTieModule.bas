Attribute VB_Name = "ApuriskBowTieModule"
Option Explicit

Public Sub Apurisk_ShowRbsTreePlaceholder()
    Apurisk_SetLastAction "Apurisk_ShowRbsTreePlaceholder"
    MsgBox "Aqui abriremos la vista de arbol RBS para navegar por categorias y riesgos.", vbInformation, "Apurisk - Analisis BowTie"
End Sub

Public Sub Apurisk_ShowBowTiePlaceholder()
    Apurisk_SetLastAction "Apurisk_ShowBowTiePlaceholder"
    MsgBox "Aqui abriremos la vista BowTie del riesgo seleccionado.", vbInformation, "Apurisk - Analisis BowTie"
End Sub

Public Sub Apurisk_ValidateCurrentContext()
    Apurisk_SetLastAction "Apurisk_ValidateCurrentContext"
    MsgBox "Aqui validaremos columnas obligatorias, RBS valido y riesgos sin clasificar.", vbInformation, "Apurisk - Validacion"
End Sub

Public Sub Apurisk_InsertBowTieValuesPlaceholder()
    Apurisk_SetLastAction "Apurisk_InsertBowTieValuesPlaceholder"
    MsgBox "Aqui escribiremos el resultado del BowTie en la tabla maestra configurada.", vbInformation, "Apurisk - Tabla maestra"
End Sub
