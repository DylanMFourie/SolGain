VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTiltOrientationPTC 
   Caption         =   "PTC Tilt and Orientation"
   ClientHeight    =   6948
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10608
   OleObjectBlob   =   "UserFormTiltOrientationPTC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormTiltOrientationPTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim collInWS As Worksheet

Private Sub CommandButtonNext_Click()

    If IsNumeric(TextBoxTilt) = False Then
        MsgBox "Please enter valid Collector Tilt"
        Exit Sub
    ElseIf IsNumeric(TextBoxOrientation) = False Then
        MsgBox "Please enter valid Collector Orientation"
        Exit Sub
    End If
    
    collInWS.Range("A2") = Val(TextBoxTilt)
    
    If Sgn(geoInWS.Range("B2")) = 1 Then
        If TextBoxOrientation.Value <= 180 Then
            collInWS.Range("B2") = -180 + TextBoxOrientation.Value
        Else
            collInWS.Range("B2") = TextBoxOrientation.Value - 180
        End If
    Else
        If TextBoxOrientation.Value <= 180 Then
            collInWS.Range("B2") = -TextBoxOrientation.Value
        Else
            collInWS.Range("B2") = 360 - TextBoxOrientation.Value
        End If
    End If
    
    UserFormTiltOrientationPTC.Hide
    UserFormConCollLayout.Show
End Sub

Private Sub CommandButtonPrevious_Click()
    Unload UserFormTiltOrientationPTC
    UserFormManualCollInputs.Show
End Sub

Private Sub UserForm_Initialize()
    Set collInWS = ThisWorkbook.Sheets("Collector Inputs")
End Sub
