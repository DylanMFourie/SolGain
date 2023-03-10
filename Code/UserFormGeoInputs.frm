VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormGeoInputs 
   Caption         =   "Geographical Inputs"
   ClientHeight    =   3996
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6780
   OleObjectBlob   =   "UserFormGeoInputs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormGeoInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim GeoInputWS As Worksheet

Private Sub CommandButtonNext_Click()
    Dim TimeZoneLongitude As Double
    
    GeoInputWS.Range("A2") = TextBoxLocation.Value
    
    If IsNumeric(TextBoxLat) = False Or Sgn(Val(TextBoxLat)) = -1 Then
        MsgBox "Enter valid latitude"
        Exit Sub
    ElseIf ComboBoxLat = "" Then
        MsgBox "Select N/S"
        Exit Sub
    ElseIf IsNumeric(TextBoxLongi) = False Or Sgn(Val(TextBoxLongi)) = -1 Then
        MsgBox "Enter valid longitude"
        Exit Sub
    ElseIf ComboBoxLongi = "" Then
        MsgBox "Select E/W"
        Exit Sub
    ElseIf ComboBoxSignTimeZone = "" Then
        MsgBox "Select +/- timezone"
        Exit Sub
    ElseIf ComboBoxTimeZoneHour = "" Then
        MsgBox "Select valid timezone"
        Exit Sub
    ElseIf ComboBoxTimeZoneMinutes = "" Then
        MsgBox "Select valid timezone"
        Exit Sub
    End If
    
    If ComboBoxLat = "N" Then
        GeoInputWS.Range("B2") = TextBoxLat.Value
    ElseIf ComboBoxLat = "S" Then
        GeoInputWS.Range("B2") = -TextBoxLat.Value
    End If
    
    If ComboBoxLongi = "E" Then
        GeoInputWS.Range("D2") = TextBoxLongi.Value
    ElseIf ComboBoxLongi = "W" Then
        GeoInputWS.Range("D2") = -TextBoxLongi.Value
    End If
    
    If ComboBoxSignTimeZone = "-" Then
        TimeZoneLongitude = 15 * CDbl(ComboBoxTimeZoneHour) + (15 / 60) * CDbl(ComboBoxTimeZoneMinutes)
    ElseIf ComboBoxSignTimeZone = "+" Then
        TimeZoneLongitude = -(15 * CDbl(ComboBoxTimeZoneHour) + (15 / 60) * CDbl(ComboBoxTimeZoneMinutes))
    End If
    
    GeoInputWS.Range("E2") = TimeZoneLongitude
    
    UserFormGeoInputs.Hide
    UserFormProcessInfo.Show
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Set GeoInputWS = ThisWorkbook.Sheets("Geographical Inputs")
    
    LabelDegree1.Caption = Chr(176)
    With ComboBoxLat
        .AddItem "N"
        .AddItem "S"
    End With
    
    LabelDegree2.Caption = Chr(176)
    With ComboBoxLongi
        .AddItem "E"
        .AddItem "W"
    End With
    
    ComboBoxSignTimeZone.AddItem "+"
    ComboBoxSignTimeZone.AddItem "-"
    
    With ComboBoxTimeZoneHour
        For i = 0 To 12
            If Abs(i) >= 10 Then
                .AddItem i
            Else
                .AddItem "0" & i
            End If
        Next i
    End With
    
    With ComboBoxTimeZoneMinutes
        .AddItem "00"
        .AddItem "15"
        .AddItem "30"
        .AddItem "45"
    End With
End Sub
