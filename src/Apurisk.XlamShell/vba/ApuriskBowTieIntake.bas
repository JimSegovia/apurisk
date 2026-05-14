Attribute VB_Name = "ApuriskBowTieIntake"
Option Explicit

Public Sub Apurisk_StartBowTieIntake()
    Dim targetWorkbook As Workbook

    Set targetWorkbook = ActiveWorkbook
    If targetWorkbook Is Nothing Then
        MsgBox "No hay un libro activo para trabajar.", vbExclamation, APURISK_DIALOG_TITLE
        Exit Sub
    End If

    Apurisk_EnsureWorkbookBase False
    frmApuriskBowTieIntake.Show vbModeless
    Apurisk_SetLastAction "Apurisk_StartBowTieIntake"
End Sub

Public Function Apurisk_PickRange(ByVal promptText As String) As String
    Dim selectedRange As Range

    On Error Resume Next
    Set selectedRange = Application.InputBox(prompt:=promptText, Title:=APURISK_DIALOG_TITLE, Type:=8)
    On Error GoTo 0

    If selectedRange Is Nothing Then
        Exit Function
    End If

    selectedRange.Interior.Color = RGB(241, 247, 191)
    Apurisk_PickRange = selectedRange.Address(External:=True)
End Function

Public Function Apurisk_RangeFromAddress(ByVal addressText As String) As Range
    If Len(Trim$(addressText)) = 0 Then
        Exit Function
    End If

    On Error Resume Next
    Set Apurisk_RangeFromAddress = Range(addressText)
    On Error GoTo 0
End Function

Public Function Apurisk_ReadConfigValue(ByVal keyName As String) As String
    Dim configSheet As Worksheet
    Dim targetRow As Long

    If ActiveWorkbook Is Nothing Then
        Exit Function
    End If

    Set configSheet = ActiveWorkbook.Worksheets(APURISK_SHEET_CONFIG)
    targetRow = Apurisk_FindConfigRow(configSheet, keyName)
    If targetRow = 0 Then
        Exit Function
    End If

    Apurisk_ReadConfigValue = Trim$(CStr(configSheet.Cells(targetRow, 2).Value))
End Function

Public Sub Apurisk_WriteConfigValue(ByVal keyName As String, ByVal keyValue As String, ByVal notes As String)
    Dim configSheet As Worksheet
    Dim targetRow As Long

    If ActiveWorkbook Is Nothing Then
        Exit Sub
    End If

    Set configSheet = ActiveWorkbook.Worksheets(APURISK_SHEET_CONFIG)
    targetRow = Apurisk_FindConfigRow(configSheet, keyName)

    If targetRow = 0 Then
        targetRow = configSheet.Cells(configSheet.Rows.Count, 1).End(xlUp).Row + 1
        If targetRow < 2 Then
            targetRow = 2
        End If
    End If

    configSheet.Cells(targetRow, 1).Value = keyName
    configSheet.Cells(targetRow, 2).Value = keyValue
    configSheet.Cells(targetRow, 3).Value = notes
    configSheet.Columns.AutoFit
End Sub

Public Function Apurisk_RequiredFieldKeys() As Variant
    Apurisk_RequiredFieldKeys = Array( _
        "RbsNameRange", "RbsCodeRange", "RiskTableRange", "RiskIdRange", "RiskTopRange", _
        "RiskRbsCodeRange", "RiskDescriptionRange", "RiskCauseRange", "RiskPotentialEffectRange", _
        "RiskProbabilityRange", "RiskImpactRange", "RiskSeverityRange", "RiskMitigationRange", _
        "RiskOwnerRange")
End Function

Public Function Apurisk_FieldLabel(ByVal fieldKey As String) As String
    Select Case fieldKey
        Case "RbsNameRange": Apurisk_FieldLabel = "Nombre RBS"
        Case "RbsCodeRange": Apurisk_FieldLabel = "Codigo RBS"
        Case "RiskTableRange": Apurisk_FieldLabel = "Seleccion automatica"
        Case "RiskIdRange": Apurisk_FieldLabel = "ID"
        Case "RiskTopRange": Apurisk_FieldLabel = "TOP"
        Case "RiskRbsCodeRange": Apurisk_FieldLabel = "Codigo RBS del riesgo"
        Case "RiskRbsNameRange": Apurisk_FieldLabel = "Nombre RBS del riesgo"
        Case "RiskDescriptionRange": Apurisk_FieldLabel = "Descripcion del riesgo"
        Case "RiskCauseRange": Apurisk_FieldLabel = "Causas clave"
        Case "RiskPotentialEffectRange": Apurisk_FieldLabel = "Impacto / efecto potencial"
        Case "RiskProbabilityRange": Apurisk_FieldLabel = "Probabilidad"
        Case "RiskImpactRange": Apurisk_FieldLabel = "Impacto"
        Case "RiskSeverityRange": Apurisk_FieldLabel = "Gravedad"
        Case "RiskMitigationRange": Apurisk_FieldLabel = "Medidas de mitigacion"
        Case "RiskOwnerRange": Apurisk_FieldLabel = "Persona responsable"
        Case Else
            If Left$(fieldKey, 14) = "ImpactCategory" Then
                Apurisk_FieldLabel = "Cat. Impacto " & Replace(fieldKey, "ImpactCategory", "")
            End If
    End Select
End Function

Public Function Apurisk_FieldNotes(ByVal fieldKey As String) As String
    Apurisk_FieldNotes = "Rango guardado para " & Apurisk_FieldLabel(fieldKey)
End Function

Public Function Apurisk_ValidateRequiredFields(ByVal formObject As Object) As Boolean
    Dim fieldKey As Variant
    Dim currentValue As String

    For Each fieldKey In Apurisk_RequiredFieldKeys()
        currentValue = Trim$(CStr(CallByName(formObject, "GetFieldValue", VbMethod, CStr(fieldKey))))
        If Len(currentValue) = 0 Then
            MsgBox "Falta completar el campo obligatorio '" & Apurisk_FieldLabel(CStr(fieldKey)) & "'.", vbExclamation, APURISK_DIALOG_TITLE
            Exit Function
        End If
    Next fieldKey

    Apurisk_ValidateRequiredFields = True
End Function

Public Function Apurisk_RiskIdExists(ByVal riskIdAddress As String, ByVal riskIdValue As String) As Boolean
    Dim idRange As Range
    Dim currentCell As Range

    Set idRange = Apurisk_RangeFromAddress(riskIdAddress)
    If idRange Is Nothing Then
        Exit Function
    End If

    For Each currentCell In idRange.Cells
        If StrComp(Trim$(CStr(currentCell.Value)), riskIdValue, vbTextCompare) = 0 Then
            Apurisk_RiskIdExists = True
            Exit Function
        End If
    Next currentCell
End Function

Public Sub Apurisk_SaveRbsSnapshot(ByVal rbsNameAddress As String, ByVal rbsCodeAddress As String)
    Dim nameRange As Range
    Dim codeRange As Range
    Dim targetSheet As Worksheet
    Dim rowCount As Long
    Dim rowIndex As Long

    Set nameRange = Apurisk_RangeFromAddress(rbsNameAddress)
    Set codeRange = Apurisk_RangeFromAddress(rbsCodeAddress)

    If nameRange Is Nothing Or codeRange Is Nothing Then
        Exit Sub
    End If

    If nameRange.Rows.Count <> codeRange.Rows.Count Then
        MsgBox "Nombre RBS y Codigo RBS deben tener la misma cantidad de filas.", vbExclamation, APURISK_DIALOG_TITLE
        Exit Sub
    End If

    Set targetSheet = ActiveWorkbook.Worksheets(APURISK_SHEET_RBS)
    targetSheet.Cells.Clear
    targetSheet.Range("A1:E1").Value = Array("CodigoRBS", "Nombre", "PadreRBS", "Nivel", "Descripcion")
    targetSheet.Rows(1).Font.Bold = True

    rowCount = nameRange.Rows.Count
    For rowIndex = 1 To rowCount
        targetSheet.Cells(rowIndex + 1, 1).Value = codeRange.Cells(rowIndex, 1).Value
        targetSheet.Cells(rowIndex + 1, 2).Value = nameRange.Cells(rowIndex, 1).Value
        targetSheet.Cells(rowIndex + 1, 3).Value = Apurisk_ParentRbsCode(CStr(codeRange.Cells(rowIndex, 1).Value))
        targetSheet.Cells(rowIndex + 1, 4).Value = Apurisk_RbsLevel(CStr(codeRange.Cells(rowIndex, 1).Value))
    Next rowIndex

    targetSheet.Columns.AutoFit
End Sub

Public Sub Apurisk_SaveMappingSnapshot(ByVal formObject As Object)
    Dim mapSheet As Worksheet
    Dim nextRow As Long
    Dim fieldKey As Variant
    Dim impactIndex As Long
    Dim impactCount As Long
    Dim currentValue As String

    Set mapSheet = ActiveWorkbook.Worksheets(APURISK_SHEET_MAP)
    mapSheet.Cells.Clear
    mapSheet.Range("A1:D1").Value = Array("CampoApurisk", "RangoExcel", "Obligatorio", "Notas")
    mapSheet.Rows(1).Font.Bold = True

    nextRow = 2
    For Each fieldKey In Array( _
        "RbsNameRange", "RbsCodeRange", "RiskTableRange", "RiskIdRange", "RiskTopRange", _
        "RiskRbsCodeRange", "RiskRbsNameRange", "RiskDescriptionRange", "RiskCauseRange", _
        "RiskPotentialEffectRange", "RiskProbabilityRange", "RiskImpactRange", "RiskSeverityRange", _
        "RiskMitigationRange", "RiskOwnerRange")
        currentValue = Trim$(CStr(CallByName(formObject, "GetFieldValue", VbMethod, CStr(fieldKey))))
        mapSheet.Cells(nextRow, 1).Value = Apurisk_FieldLabel(CStr(fieldKey))
        mapSheet.Cells(nextRow, 2).Value = currentValue
        mapSheet.Cells(nextRow, 3).Value = IIf(Apurisk_IsRequiredField(CStr(fieldKey)), "Si", "No")
        mapSheet.Cells(nextRow, 4).Value = Apurisk_FieldNotes(CStr(fieldKey))
        Apurisk_WriteConfigValue "Field." & CStr(fieldKey), currentValue, Apurisk_FieldNotes(CStr(fieldKey))
        nextRow = nextRow + 1
    Next fieldKey

    impactCount = Apurisk_GetImpactFieldCount()
    For impactIndex = 1 To impactCount
        currentValue = Trim$(CStr(CallByName(formObject, "GetFieldValue", VbMethod, "ImpactCategory" & impactIndex)))
        mapSheet.Cells(nextRow, 1).Value = "Cat. Impacto " & impactIndex
        mapSheet.Cells(nextRow, 2).Value = currentValue
        mapSheet.Cells(nextRow, 3).Value = "No"
        mapSheet.Cells(nextRow, 4).Value = Apurisk_FieldNotes("ImpactCategory" & impactIndex)
        Apurisk_WriteConfigValue "Field.ImpactCategory" & impactIndex, currentValue, Apurisk_FieldNotes("ImpactCategory" & impactIndex)
        nextRow = nextRow + 1
    Next impactIndex

    mapSheet.Columns.AutoFit
End Sub

Public Function Apurisk_GetImpactFieldCount() As Long
    Apurisk_GetImpactFieldCount = Val(Apurisk_ReadConfigValue("ImpactFieldCount"))
    If Apurisk_GetImpactFieldCount < 1 Then
        Apurisk_GetImpactFieldCount = 1
    End If
End Function

Public Sub Apurisk_SetImpactFieldCount(ByVal impactCount As Long)
    Apurisk_WriteConfigValue "ImpactFieldCount", CStr(impactCount), "Cantidad de impactos configurables en el popup"
End Sub

Public Function Apurisk_IsRequiredField(ByVal fieldKey As String) As Boolean
    Dim requiredField As Variant

    For Each requiredField In Apurisk_RequiredFieldKeys()
        If StrComp(CStr(requiredField), fieldKey, vbTextCompare) = 0 Then
            Apurisk_IsRequiredField = True
            Exit Function
        End If
    Next requiredField
End Function

Private Function Apurisk_ParentRbsCode(ByVal rbsCode As String) As String
    Dim lastDot As Long

    lastDot = InStrRev(rbsCode, ".")
    If lastDot > 0 Then
        Apurisk_ParentRbsCode = Left$(rbsCode, lastDot - 1)
    End If
End Function

Private Function Apurisk_RbsLevel(ByVal rbsCode As String) As Long
    If Len(Trim$(rbsCode)) = 0 Then
        Exit Function
    End If

    Apurisk_RbsLevel = UBound(Split(rbsCode, ".")) + 1
End Function

Private Function Apurisk_FindConfigRow(ByVal configSheet As Worksheet, ByVal keyName As String) As Long
    Dim rowIndex As Long
    Dim lastRow As Long

    lastRow = configSheet.Cells(configSheet.Rows.Count, 1).End(xlUp).Row
    For rowIndex = 2 To lastRow
        If StrComp(Trim$(CStr(configSheet.Cells(rowIndex, 1).Value)), keyName, vbTextCompare) = 0 Then
            Apurisk_FindConfigRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function
