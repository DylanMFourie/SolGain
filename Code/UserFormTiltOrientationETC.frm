VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTiltOrientationETC 
   Caption         =   "ETC Tilt and Orientation"
   ClientHeight    =   6444
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13416
   OleObjectBlob   =   "UserFormTiltOrientationETC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormTiltOrientationETC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim collInWS As Worksheet

Private Sub CommandButtonNext_Click()
    
    If IsNumeric(TextBoxTilt) = False Or Sgn(TextBoxTilt) = -1 Then
        MsgBox "Please enter valid Collector Tilt"
        Exit Sub
    ElseIf IsNumeric(TextBoxOrientation) = False Or Sgn(TextBoxOrientation) = -1 Then
        MsgBox "Please enter valid Collector Orientation"
        Exit Sub
    ElseIf ComboBoxTubes = "" Then
        MsgBox "Please select Tube Configuration"
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
    
    If ComboBoxTubes = "Horizontal" Then
        collInWS.Range("C2") = False
    ElseIf ComboBoxTubes = "Vertical" Then
        collInWS.Range("C2") = True
    End If
    
    UserFormTiltOrientationETC.Hide
    UserFormCollectorLayout.Show
End Sub

Private Sub CommandButtonPrevious_Click()
    Unload UserFormTiltOrientationETC
    UserFormManualCollInputs.Show
End Sub

Private Sub UserForm_Initialize()

    Set collInWS = ThisWorkbook.Sheets("Collector Inputs")
    
    ComboBoxTubes.AddItem "Vertical"
    ComboBoxTubes.AddItem "Horizontal"

End Sub
