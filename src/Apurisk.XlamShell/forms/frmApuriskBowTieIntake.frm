VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmApuriskBowTieIntake
   Caption         =   "Apurisk - Ingresar Valores BowTie"
   ClientHeight    =   4680
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8550
   OleObjectBlob   =   "frmApuriskBowTieIntake.frx":0000
   StartUpPosition = 1  'Centrar en propietario
End
Attribute VB_Name = "frmApuriskBowTieIntake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ActiveFieldKey As String
Private m_ActiveTextBoxName As String
Private m_ImpactCount As Long

Private Sub UserForm_Initialize()
    m_ImpactCount = Apurisk_GetImpactFieldCount()
    RenderImpactFields
    LoadSavedValues
    SetActiveField "RbsNameRange", "txtRbsNameRange"
End Sub

Private Sub btnLoadRange_Click()
    Dim pickedAddress As String

    If Len(m_ActiveFieldKey) = 0 Then
        MsgBox "Selecciona primero un cuadro del popup.", vbInformation, APURISK_DIALOG_TITLE
        Exit Sub
    End If

    Me.Hide
    pickedAddress = Apurisk_PickRange("Selecciona el rango para '" & Apurisk_FieldLabel(m_ActiveFieldKey) & "'.")
    Me.Show vbModeless

    If Len(pickedAddress) = 0 Then
        Exit Sub
    End If

    SetFieldValue m_ActiveFieldKey, pickedAddress
End Sub

Private Sub btnAddImpact_Click()
    m_ImpactCount = m_ImpactCount + 1
    Apurisk_SetImpactFieldCount m_ImpactCount
    RenderImpactFields
    LoadSavedValues
    SetActiveField "ImpactCategory" & m_ImpactCount, "txtImpactCategory" & m_ImpactCount
End Sub

Private Sub btnAceptar_Click()
    Dim selectedRiskId As String

    If Not Apurisk_ValidateRequiredFields(Me) Then
        Exit Sub
    End If

    selectedRiskId = Trim$(txtSelectedRiskId.Text)
    If Len(selectedRiskId) = 0 Then
        MsgBox "Ingresa el ID del riesgo a analizar antes de aceptar.", vbExclamation, APURISK_DIALOG_TITLE
        Exit Sub
    End If

    If Not Apurisk_RiskIdExists(GetFieldValue("RiskIdRange"), selectedRiskId) Then
        MsgBox "El ID '" & selectedRiskId & "' no existe dentro del rango seleccionado.", vbExclamation, APURISK_DIALOG_TITLE
        Exit Sub
    End If

    Apurisk_WriteConfigValue "SelectedRiskId", selectedRiskId, "ID elegido para iniciar el analisis BowTie"
    Apurisk_SaveRbsSnapshot GetFieldValue("RbsNameRange"), GetFieldValue("RbsCodeRange")
    Apurisk_SaveMappingSnapshot Me
    Apurisk_SetLastAction "frmApuriskBowTieIntake.Accept"

    MsgBox "Los rangos quedaron guardados y se mantendran cuando vuelvas a abrir esta ventana.", vbInformation, APURISK_DIALOG_TITLE
    Me.Hide
End Sub

Private Sub btnCancelar_Click()
    Apurisk_SetLastAction "frmApuriskBowTieIntake.Cancel"
    Me.Hide
End Sub

Public Function GetFieldValue(ByVal fieldKey As String) As String
    Select Case fieldKey
        Case "RbsNameRange": GetFieldValue = txtRbsNameRange.Text
        Case "RbsCodeRange": GetFieldValue = txtRbsCodeRange.Text
        Case "RiskTableRange": GetFieldValue = txtRiskTableRange.Text
        Case "RiskIdRange": GetFieldValue = txtRiskIdRange.Text
        Case "RiskTopRange": GetFieldValue = txtRiskTopRange.Text
        Case "RiskRbsCodeRange": GetFieldValue = txtRiskRbsCodeRange.Text
        Case "RiskRbsNameRange": GetFieldValue = txtRiskRbsNameRange.Text
        Case "RiskDescriptionRange": GetFieldValue = txtRiskDescriptionRange.Text
        Case "RiskCauseRange": GetFieldValue = txtRiskCauseRange.Text
        Case "RiskPotentialEffectRange": GetFieldValue = txtRiskPotentialEffectRange.Text
        Case "RiskProbabilityRange": GetFieldValue = txtRiskProbabilityRange.Text
        Case "RiskImpactRange": GetFieldValue = txtRiskImpactRange.Text
        Case "RiskSeverityRange": GetFieldValue = txtRiskSeverityRange.Text
        Case "RiskMitigationRange": GetFieldValue = txtRiskMitigationRange.Text
        Case "RiskOwnerRange": GetFieldValue = txtRiskOwnerRange.Text
        Case Else
            If Left$(fieldKey, 14) = "ImpactCategory" Then
                GetFieldValue = Me.Controls("txt" & fieldKey).Text
            End If
    End Select
End Function

Private Sub SetFieldValue(ByVal fieldKey As String, ByVal fieldValue As String)
    Select Case fieldKey
        Case "RbsNameRange": txtRbsNameRange.Text = fieldValue
        Case "RbsCodeRange": txtRbsCodeRange.Text = fieldValue
        Case "RiskTableRange": txtRiskTableRange.Text = fieldValue
        Case "RiskIdRange": txtRiskIdRange.Text = fieldValue
        Case "RiskTopRange": txtRiskTopRange.Text = fieldValue
        Case "RiskRbsCodeRange": txtRiskRbsCodeRange.Text = fieldValue
        Case "RiskRbsNameRange": txtRiskRbsNameRange.Text = fieldValue
        Case "RiskDescriptionRange": txtRiskDescriptionRange.Text = fieldValue
        Case "RiskCauseRange": txtRiskCauseRange.Text = fieldValue
        Case "RiskPotentialEffectRange": txtRiskPotentialEffectRange.Text = fieldValue
        Case "RiskProbabilityRange": txtRiskProbabilityRange.Text = fieldValue
        Case "RiskImpactRange": txtRiskImpactRange.Text = fieldValue
        Case "RiskSeverityRange": txtRiskSeverityRange.Text = fieldValue
        Case "RiskMitigationRange": txtRiskMitigationRange.Text = fieldValue
        Case "RiskOwnerRange": txtRiskOwnerRange.Text = fieldValue
        Case Else
            If Left$(fieldKey, 14) = "ImpactCategory" Then
                Me.Controls("txt" & fieldKey).Text = fieldValue
            End If
    End Select
End Sub

Private Sub LoadSavedValues()
    Dim impactIndex As Long

    txtRbsNameRange.Text = Apurisk_ReadConfigValue("Field.RbsNameRange")
    txtRbsCodeRange.Text = Apurisk_ReadConfigValue("Field.RbsCodeRange")
    txtRiskTableRange.Text = Apurisk_ReadConfigValue("Field.RiskTableRange")
    txtRiskIdRange.Text = Apurisk_ReadConfigValue("Field.RiskIdRange")
    txtRiskTopRange.Text = Apurisk_ReadConfigValue("Field.RiskTopRange")
    txtRiskRbsCodeRange.Text = Apurisk_ReadConfigValue("Field.RiskRbsCodeRange")
    txtRiskRbsNameRange.Text = Apurisk_ReadConfigValue("Field.RiskRbsNameRange")
    txtRiskDescriptionRange.Text = Apurisk_ReadConfigValue("Field.RiskDescriptionRange")
    txtRiskCauseRange.Text = Apurisk_ReadConfigValue("Field.RiskCauseRange")
    txtRiskPotentialEffectRange.Text = Apurisk_ReadConfigValue("Field.RiskPotentialEffectRange")
    txtRiskProbabilityRange.Text = Apurisk_ReadConfigValue("Field.RiskProbabilityRange")
    txtRiskImpactRange.Text = Apurisk_ReadConfigValue("Field.RiskImpactRange")
    txtRiskSeverityRange.Text = Apurisk_ReadConfigValue("Field.RiskSeverityRange")
    txtRiskMitigationRange.Text = Apurisk_ReadConfigValue("Field.RiskMitigationRange")
    txtRiskOwnerRange.Text = Apurisk_ReadConfigValue("Field.RiskOwnerRange")
    txtSelectedRiskId.Text = Apurisk_ReadConfigValue("SelectedRiskId")

    For impactIndex = 1 To m_ImpactCount
        Me.Controls("txtImpactCategory" & impactIndex).Text = Apurisk_ReadConfigValue("Field.ImpactCategory" & impactIndex)
    Next impactIndex
End Sub

Private Sub RenderImpactFields()
    Dim impactIndex As Long
    Dim baseTop As Single
    Dim labelControl As MSForms.Control
    Dim textControl As MSForms.Control

    ClearDynamicImpactControls
    baseTop = 455

    For impactIndex = 1 To m_ImpactCount
        Set labelControl = Me.Controls.Add("Forms.Label.1", "lblImpactCategory" & impactIndex, True)
        labelControl.Caption = "Cat. Impacto " & impactIndex
        labelControl.Left = 440
        labelControl.Top = baseTop + ((impactIndex - 1) * 28)
        labelControl.Width = 90
        labelControl.Height = 18
        labelControl.BackStyle = 0

        Set textControl = Me.Controls.Add("Forms.TextBox.1", "txtImpactCategory" & impactIndex, True)
        textControl.Left = 530
        textControl.Top = baseTop + ((impactIndex - 1) * 28) - 2
        textControl.Width = 150
        textControl.Height = 20
        textControl.BackColor = 16777215
        textControl.BorderStyle = 1
    Next impactIndex

    btnAceptar.Top = baseTop + (m_ImpactCount * 28) + 12
    btnCancelar.Top = btnAceptar.Top
End Sub

Private Sub ClearDynamicImpactControls()
    Dim impactIndex As Long

    For impactIndex = Me.Controls.Count - 1 To 0 Step -1
        If TypeName(Me.Controls.Item(impactIndex)) <> "CommandButton" Then
            If Left$(Me.Controls.Item(impactIndex).Name, 17) = "lblImpactCategory" Or Left$(Me.Controls.Item(impactIndex).Name, 17) = "txtImpactCategory" Then
                Me.Controls.Remove Me.Controls.Item(impactIndex).Name
            End If
        End If
    Next impactIndex
End Sub

Private Sub SetActiveField(ByVal fieldKey As String, ByVal textBoxName As String)
    ResetFieldHighlights
    m_ActiveFieldKey = fieldKey
    m_ActiveTextBoxName = textBoxName
    Me.Controls(textBoxName).BackColor = 15787967
End Sub

Private Sub ResetFieldHighlights()
    Dim controlItem As MSForms.Control

    For Each controlItem In Me.Controls
        If TypeName(controlItem) = "TextBox" Then
            controlItem.BackColor = 16777215
        End If
    Next controlItem
End Sub

Private Sub txtRbsNameRange_Enter(): SetActiveField "RbsNameRange", "txtRbsNameRange": End Sub
Private Sub txtRbsCodeRange_Enter(): SetActiveField "RbsCodeRange", "txtRbsCodeRange": End Sub
Private Sub txtRiskTableRange_Enter(): SetActiveField "RiskTableRange", "txtRiskTableRange": End Sub
Private Sub txtRiskIdRange_Enter(): SetActiveField "RiskIdRange", "txtRiskIdRange": End Sub
Private Sub txtRiskTopRange_Enter(): SetActiveField "RiskTopRange", "txtRiskTopRange": End Sub
Private Sub txtRiskRbsCodeRange_Enter(): SetActiveField "RiskRbsCodeRange", "txtRiskRbsCodeRange": End Sub
Private Sub txtRiskRbsNameRange_Enter(): SetActiveField "RiskRbsNameRange", "txtRiskRbsNameRange": End Sub
Private Sub txtRiskDescriptionRange_Enter(): SetActiveField "RiskDescriptionRange", "txtRiskDescriptionRange": End Sub
Private Sub txtRiskCauseRange_Enter(): SetActiveField "RiskCauseRange", "txtRiskCauseRange": End Sub
Private Sub txtRiskPotentialEffectRange_Enter(): SetActiveField "RiskPotentialEffectRange", "txtRiskPotentialEffectRange": End Sub
Private Sub txtRiskProbabilityRange_Enter(): SetActiveField "RiskProbabilityRange", "txtRiskProbabilityRange": End Sub
Private Sub txtRiskImpactRange_Enter(): SetActiveField "RiskImpactRange", "txtRiskImpactRange": End Sub
Private Sub txtRiskSeverityRange_Enter(): SetActiveField "RiskSeverityRange", "txtRiskSeverityRange": End Sub
Private Sub txtRiskMitigationRange_Enter(): SetActiveField "RiskMitigationRange", "txtRiskMitigationRange": End Sub
Private Sub txtRiskOwnerRange_Enter(): SetActiveField "RiskOwnerRange", "txtRiskOwnerRange": End Sub

