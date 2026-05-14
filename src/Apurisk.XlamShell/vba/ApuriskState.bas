Attribute VB_Name = "ApuriskState"
Option Explicit

' Shared state stays small on purpose.
' The workbook is the source of truth; globals only cache session-level info.
Public g_ApuriskLoaded As Boolean
Public g_LastAction As String

Public Const APURISK_SHEET_CONFIG As String = "Apurisk_Config"
Public Const APURISK_SHEET_RBS As String = "Apurisk_RBS"
Public Const APURISK_SHEET_MAP As String = "Apurisk_RiskMaster_Map"
Public Const APURISK_SHEET_WORK As String = "Apurisk_BowTie_Work"
Public Const APURISK_SHEET_DIAGRAM As String = "Apurisk_Diagram"
Public Const APURISK_DIALOG_TITLE As String = "Analisis BowTie - Ingresar Valores"

Public Sub Apurisk_SetLastAction(ByVal actionName As String)
    g_LastAction = actionName
End Sub
