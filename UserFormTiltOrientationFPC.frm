VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTiltOrientationFPC 
   Caption         =   "Collector Tilt and Orientation"
   ClientHeight    =   6600
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10008
   OleObjectBlob   =   "UserFormTiltOrientationFPC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormTiltOrientationFPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim collInWS As Worksheet
Dim geoInWS As Worksheet

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
    
    UserFormTiltOrientationFPC.Hide
    UserFormCollectorLayout.Show
End Sub

Private Sub CommandButtonPrevious_Click()
    Unload UserFormTiltOrientationFPC
    UserFormManualCollInputs.Show
End Sub

Private Sub UserForm_Initialize()
    Set collInWS = ThisWorkbook.Sheets("Collector Inputs")
    Set geoInWS = ThisWorkbook.Sheets("Geographical Inputs")
End Sub
