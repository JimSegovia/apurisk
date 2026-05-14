Attribute VB_Name = "ApuriskBootstrap"
Option Explicit

Public Sub Auto_Open()
    Apurisk_Initialize
End Sub

Public Sub Auto_Close()
    g_ApuriskLoaded = False
    Apurisk_SetLastAction "Auto_Close"
End Sub

Public Sub Apurisk_Initialize()
    g_ApuriskLoaded = True
    Apurisk_SetLastAction "Apurisk_Initialize"
End Sub
