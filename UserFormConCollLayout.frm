VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormConCollLayout 
   Caption         =   "Collector Field Layout"
   ClientHeight    =   8268.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13224
   OleObjectBlob   =   "UserFormConCollLayout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormConCollLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim collInWS As Worksheet

Private Sub CommandButtonNext_Click()

    If IsNumeric(TextBoxNoRows) = False Then
        MsgBox "Please enter valid Number of Rows"
        Exit Sub
    ElseIf IsNumeric(TextBoxNoColumnsSeries) = False Then
        MsgBox "Please enter valid Number of Columns in Series"
        Exit Sub
    ElseIf IsNumeric(TextBoxNoModulesParallel) = False Then
        MsgBox "Please enter valid Number of Modules in Parallel"
        Exit Sub
    ElseIf IsNumeric(TextBoxDistRows) = False Then
        MsgBox "Please enter valid Distance between Rows"
    End If
    
    collInWS.Range("F2") = Val(TextBoxNoRows) * Val(TextBoxNoColumnsSeries)
    collInWS.Range("G2") = Val(TextBoxNoModulesParallel)
    
    collInWS.Range("H2") = Val(TextBoxNoRows)
    collInWS.Range("I2") = Val(TextBoxNoColumnsSeries) * Val(TextBoxNoModulesParallel)
    collInWS.Range("J2") = Val(TextBoxDistRows)
    
    UserFormConCollLayout.Hide
    UserFormMiscellaneous.Show
End Sub

Private Sub CommandButtonPrevious_Click()
    Unload UserFormConCollLayout
    UserFormTiltOrientationPTC.Show
End Sub


Private Sub UserForm_Initialize()
    Set collInWS = ThisWorkbook.Sheets("Collector Inputs")
End Sub
