VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormCollectorLayout 
   Caption         =   "Collector Field Layout"
   ClientHeight    =   7116
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11124
   OleObjectBlob   =   "UserFormCollectorLayout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormCollectorLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim collInWS As Worksheet

Private Sub CommandButtonNext_Click()
    
    If IsNumeric(TextBoxNoCollsSeries) = False Then
        MsgBox "Please enter valid Number of Collectors in Series"
        Exit Sub
    ElseIf IsNumeric(TextBoxNoModulesParallel) = False Then
        MsgBox "Please enter valid Number of Modules in Parallel"
        Exit Sub
    End If
    
    collInWS.Range("F2") = Val(TextBoxNoCollsSeries)
    collInWS.Range("G2") = Val(TextBoxNoModulesParallel)
    collInWS.Range("H2") = ""
    collInWS.Range("I2") = ""
    
    UserFormCollectorLayout.Hide
    UserFormMiscellaneous.Show
End Sub

Private Sub CommandButtonPrevious_Click()
    Unload UserFormCollectorLayout
    
    If UserFormManualCollInputs.ComboBoxCollectorType = FPC Then
        UserFormTiltOrientationFPC.Show
    ElseIf UserFormManualCollInputs.ComboBoxCollectorType = ETC Then
        UserFormTiltOrientationETC.Show
    End If

End Sub

Private Sub UserForm_Initialize()
    Set collInWS = ThisWorkbook.Sheets("Collector Inputs")
End Sub
