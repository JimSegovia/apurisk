Attribute VB_Name = "ApuriskWorkbookSetup"
Option Explicit

Public Sub Apurisk_EnsureWorkbookBase(Optional ByVal showConfirmation As Boolean = True)
    Dim targetWorkbook As Workbook

    Set targetWorkbook = ActiveWorkbook
    If targetWorkbook Is Nothing Then
        MsgBox "No hay un libro activo para preparar.", vbExclamation, "Apurisk"
        Exit Sub
    End If

    Apurisk_EnsureSheet targetWorkbook, APURISK_SHEET_CONFIG, Array("Parametro", "Valor", "Notas")
    Apurisk_EnsureSheet targetWorkbook, APURISK_SHEET_RBS, Array("CodigoRBS", "Nombre", "PadreRBS", "Nivel", "Descripcion")
    Apurisk_EnsureSheet targetWorkbook, APURISK_SHEET_MAP, Array("CampoApurisk", "RangoExcel", "Obligatorio", "Notas")
    Apurisk_EnsureSheet targetWorkbook, APURISK_SHEET_WORK, Array("RiskID", "RBS", "Elemento", "Tipo", "Valor", "Owner", "Efectividad", "Notas")
    Apurisk_EnsureSheet targetWorkbook, APURISK_SHEET_DIAGRAM, Array("Area reservada para el diagrama BowTie")

    targetWorkbook.Worksheets(APURISK_SHEET_CONFIG).Activate
    Apurisk_SetLastAction "Apurisk_EnsureWorkbookBase"

    If showConfirmation Then
        MsgBox "Base inicial creada. El siguiente paso sera configurar tabla maestra y catalogo RBS.", vbInformation, "Apurisk"
    End If
End Sub

Private Sub Apurisk_EnsureSheet(ByVal targetWorkbook As Workbook, ByVal sheetName As String, ByVal headers As Variant)
    Dim targetSheet As Worksheet
    Dim headerIndex As Long

    Set targetSheet = Apurisk_FindSheet(targetWorkbook, sheetName)
    If targetSheet Is Nothing Then
        Set targetSheet = targetWorkbook.Worksheets.Add(After:=targetWorkbook.Worksheets(targetWorkbook.Worksheets.Count))
        targetSheet.Name = sheetName
    End If

    For headerIndex = LBound(headers) To UBound(headers)
        targetSheet.Cells(1, headerIndex + 1).Value = headers(headerIndex)
        targetSheet.Cells(1, headerIndex + 1).Font.Bold = True
    Next headerIndex

    targetSheet.Rows(1).Interior.Color = RGB(220, 230, 241)
    targetSheet.Columns.AutoFit
End Sub

Private Function Apurisk_FindSheet(ByVal targetWorkbook As Workbook, ByVal sheetName As String) As Worksheet
    Dim currentSheet As Worksheet

    For Each currentSheet In targetWorkbook.Worksheets
        If StrComp(currentSheet.Name, sheetName, vbTextCompare) = 0 Then
            Set Apurisk_FindSheet = currentSheet
            Exit Function
        End If
    Next currentSheet
End Function
