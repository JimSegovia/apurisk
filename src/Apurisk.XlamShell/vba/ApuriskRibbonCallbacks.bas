Attribute VB_Name = "ApuriskRibbonCallbacks"
Option Explicit

' Ribbon callbacks stay thin and forward the work to module-level procedures.
Public Sub Apurisk_OpenBowTieIntake(control As IRibbonControl)
    Apurisk_StartBowTieIntake
End Sub

Public Sub Apurisk_OpenRbsTree(control As IRibbonControl)
    Apurisk_ShowRbsTreePlaceholder
End Sub

Public Sub Apurisk_OpenBowTie(control As IRibbonControl)
    Apurisk_ShowBowTiePlaceholder
End Sub

Public Sub Apurisk_ValidateSelection(control As IRibbonControl)
    Apurisk_ValidateCurrentContext
End Sub

Public Sub Apurisk_InsertValues(control As IRibbonControl)
    Apurisk_InsertBowTieValuesPlaceholder
End Sub
