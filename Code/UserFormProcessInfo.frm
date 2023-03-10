VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormProcessInfo 
   Caption         =   "Process"
   ClientHeight    =   5376
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10092
   OleObjectBlob   =   "UserFormProcessInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormProcessInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim proInWS As Worksheet
Dim collInWS As Worksheet

Private Sub CommandButtonNext_Click()

    If IsNumeric(TextBoxFeedTemp.Value) = False Then
        MsgBox "Please enter valid feed temperature"
        Exit Sub
    ElseIf IsNumeric(TextBoxReturnTemp.Value) = False Then
        MsgBox "Please enter valid return temperature"
        Exit Sub
    ElseIf IsNumeric(TextBoxMassFlow.Value) = False Then
        MsgBox "Please enter mass flow"
        Exit Sub
    ElseIf TextBoxMassFlow.Value < 1 Then
        MsgBox "Please note Mass Flow Rate units are kg/h"
    ElseIf IsNumeric(TextBoxHeatCapacity.Value) = False Then
        MsgBox "Please enter heat capacity"
        Exit Sub
    End If
    
    proInWS.Range("A2") = Val(TextBoxFeedTemp)
    proInWS.Range("B2") = Val(TextBoxReturnTemp)
    proInWS.Range("C2") = Val(TextBoxHeatCapacity)
    proInWS.Range("D2") = Val(TextBoxMassFlow)
    proInWS.Range("E2") = Val(TextBoxDensity)
    
    UserFormProcessInfo.Hide
    UserFormDailyDemand.Show
End Sub

Private Sub CommandButtonPrev_Click()
    Unload UserFormProcessInfo
    UserFormTiltOrientation.Show
End Sub

Private Sub UserForm_Initialize()
    Set proInWS = ThisWorkbook.Sheets("Process Inputs")
    Set collInWS = ThisWorkbook.Sheets("Collector Inputs")
    
    LabelFeedTemp.Caption = Chr(176) & "C"
    LabelReturnTemp.Caption = Chr(176) & "C"
    LabelMassFlow.Caption = "kg/h"
    LabelHeatCap.Caption = "kJ/(kg*K)"
    LabelDensity.Caption = "kg/m^3"
    LabelNote.Caption = "*Only neccessary if process medium is not water (water: heat capacity = 4.18 kJ/(kg*K) and density = 1000 kg/m^3)"
    
    TextBoxHeatCapacity.Text = "4.18"
    TextBoxDensity.Text = "1000"
End Sub
