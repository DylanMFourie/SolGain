VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormFirstCollector 
   Caption         =   "Solar Collector Properties"
   ClientHeight    =   4608
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8208.001
   OleObjectBlob   =   "UserFormFirstCollector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormFirstCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim collInWS As Worksheet
Dim weatherWS As Worksheet
Dim geoWS As Worksheet

Private Sub ButtonAuto_Click()
    TextBoxAvailArea.Enabled = True
    TextBoxAvailArea.BackColor = White
End Sub

Private Sub ButtonManual_Click()
    TextBoxAvailArea.Value = ""
    TextBoxAvailArea.Enabled = False
    TextBoxAvailArea.BackColor = Grey
End Sub

'Public ProHeat As String
'Public DirectSteam As String
'Public FPC As String
'Public ETC As String
'Public CPC As String
'Public Water As String
'Public Glycol20 As String
'Public Glycol33 As String
'Public Glycol43 As String = "Glycol 43% (Outside Temps above -25?C)"
'Public   Glycol54 As String = "Glycol 54% (Outside Temps above -40?C)"
'Public   ParabolicTrough As String = "Parabolic Trough"
'Public   SolarOil As String = "Solar Oil"

Private Sub CommandButtonNext_Click()

    If ButtonManual.Value = True Then
        UserFormFirstCollector.Hide
        collInWS.Range("A5") = True
        collInWS.Range("B5") = ""
        UserFormManualCollInputs.Show
    ElseIf ButtonAuto.Value = True Then
        If IsNumeric(TextBoxAvailArea.Value) = False Or Sgn(TextBoxAvailArea.Value) = -1 Then
            MsgBox "Enter valid available area"
            Exit Sub
        Else
            UserFormFirstCollector.Hide
            collInWS.Range("A5") = False
            collInWS.Range("F2") = ""
            collInWS.Range("G2") = ""
            collInWS.Range("H2") = ""
            collInWS.Range("I2") = ""
            collInWS.Range("B5") = TextBoxAvailArea.Value
            Call autoAssign
            UserFormMiscellaneous.Show
        End If
    Else
        MsgBox "Please select an option"
    End If
End Sub

Private Sub CommandButtonPrev_Click()
    Unload UserFormFirstCollector
    UserFormYearlyDemand.Show
End Sub


Private Sub UserForm_Initialize()
    Set collInWS = ThisWorkbook.Sheets("Collector Inputs")
    Set weatherWS = ThisWorkbook.Sheets("Weather Data")
    Set geoWS = ThisWorkbook.Sheets("Geographical Inputs")
    
    TextBoxAvailArea.Enabled = False
    TextBoxAvailArea.BackColor = Grey
    
'    ProHeat = "Non-Concentrating (Temps less than ~110?C)"
'    DirectSteam = "Concentrating (Direct steam generation T > 110?C)"
'    FPC = "Flate Plate Collector (T < 60degC)"
'    ETC = "Evacuated Tube Collector (T < 100degC)"
'    CPC = "Compound Parabolic Collector (T < 100degC)"
'    Water = "Water (Outside Temps above 0?C)"
'    Glycol20 = "Glycol 20% (Outside Temps above -6?C)"
'    Glycol33 = "Glycol 33% (Outside Temps above -15?C)"
End Sub

Private Sub autoAssign()
    Dim theta_t() As Variant
    Dim K_t() As Variant
    Dim theta_l() As Variant
    Dim K_l() As Variant
    Dim i As Integer
    
    collInWS.Range("K2", "AA10").Clear
    
    If UserFormProcessInfo.TextBoxFeedTemp <= 80 Then
        collInWS.Range("A2") = Abs(geoWS.Range("B2"))
        collInWS.Range("B2") = 0
        
        collInWS.Range("K2") = False
        collInWS.Range("L2") = FPC
        
        'Input values of S.O.L.I.D. Flat Plat Collector (SolarKeymark license no. 011-7S839 F)
        collInWS.Range("M2") = 3.85
        collInWS.Range("N2") = 2.05
        collInWS.Range("O2") = 2.076
        collInWS.Range("P2") = ""
        collInWS.Range("Q2") = 0.811
        collInWS.Range("R2") = 2.71
        collInWS.Range("S2") = 0.01
        collInWS.Range("T2") = 7050
        collInWS.Range("U2") = 0.912
        
        collInWS.Range("V2") = 1
        collInWS.Range("W2") = 50
        collInWS.Range("X2") = 0.96
        
        collInWS.Range("Y2") = 1
        collInWS.Range("Z2") = 50
        collInWS.Range("AA2") = 0.96
        
    ElseIf UserFormProcessInfo.TextBoxFeedTemp <= 110 Then
        collInWS.Range("A2") = Abs(geoWS.Range("B2"))
        collInWS.Range("B2") = 0
        collInWS.Range("C2") = True
        
        collInWS.Range("K2") = False
        collInWS.Range("L2") = ETC
        
        'Input values of Ritter Energie ETC (SolarKeymark lisence no. 011-7S1889 R)
        collInWS.Range("M2") = 3#
        collInWS.Range("N2") = 2.058
        collInWS.Range("O2") = 1.628
        collInWS.Range("P2") = ""
        collInWS.Range("Q2") = 0.687
        collInWS.Range("R2") = 0.613
        collInWS.Range("S2") = 0.003
        collInWS.Range("T2") = 8780
        collInWS.Range("U2") = 0.912
        
        'Transverse IAMs
        collInWS.Range("V2") = 6
        collInWS.Range("Y2") = 7
        
        theta_t = Array(10, 20, 30, 40, 60, 70)
        K_t = Array(1.01, 1.02, 1.02, 1.02, 1.06, 1.2)
        theta_l = Array(10, 20, 30, 40, 50, 60, 70)
        K_l = Array(1#, 0.99, 0.97, 0.94, 0.9, 0.86, 0.85)
        
        For i = 0 To 6
            If i < 6 Then
                collInWS.Range("W" & i + 2) = theta_t(i)
                collInWS.Range("X" & i + 2) = K_t(i)
            End If
            collInWS.Range("Z" & i + 2) = theta_l(i)
            collInWS.Range("AA" & i + 2) = K_l(i)
        Next i
    Else
        collInWS.Range("A2") = 0
        collInWS.Range("B2") = 0
        
        collInWS.Range("J2") = 0
        collInWS.Range("K2") = True
        collInWS.Range("L2") = ParabolicTrough
        'Input values of NEP Solar Polytrough 1800
        collInWS.Range("M2") = 9.225
        collInWS.Range("N2") = 5
        collInWS.Range("O2") = 1.845
        collInWS.Range("P2") = 0.65
        collInWS.Range("Q2") = 0.689
        collInWS.Range("R2") = 0.36
        collInWS.Range("S2") = 0.0011
        collInWS.Range("T2") = 5224.932
        collInWS.Range("U2") = 0.912
        collInWS.Range("V2") = 0
        collInWS.Range("W2", "X10") = ""
        
        collInWS.Range("Y2") = 8
        theta_l = Array(10#, 20#, 30#, 40#, 50#, 60#, 70#, 80#)
        K_l = Array(0.99, 0.99, 0.98, 0.96, 0.93, 0.88, 0.75, 0.46)
        
        For i = 0 To 7
            collInWS.Range("Z" & i + 2) = theta_l(i)
            collInWS.Range("AA" & i + 2) = K_l(i)
        Next i
        
    End If
    
    Call assignColHTF
End Sub

Public Sub assignColHTF()
'    Public Const Water As String = "Water (Outside Temps above 0?C)"
'    Public Const Glycol20 As String = "Glycol 20% (Outside Temps above -6?C)"
'    Public Const Glycol33 As String = "Glycol 33% (Outside Temps above -15?C)"
'    Public Const Glycol43 As String = "Glycol 43% (Outside Temps above -25?C)"
'    Public Const Glycol54 As String = "Glycol 54% (Outside Temps above -40?C)"
'    Public Const SolarOil As String = "Solar Oil"
    Dim Cp_water As Double
    Dim Cp_glycol As Double
    
    Cp_water = 4.18 'kJ/(kg*K)
    Cp_glycol = 2.5 'kJ/(kg*K)
    
    With collInWS
        If .Range("K2") = False Then
            If WorksheetFunction.Min(weatherWS.Range("F4", "F8764")) >= 0 Then
                .Range("E2") = Water
                .Range("C5") = Cp_water
                .Range("D5") = 1000
            ElseIf WorksheetFunction.Min(weatherWS.Range("F4", "F8764")) >= -10.6 Then
                .Range("E2") = Glycol20
                .Range("C5") = 0.8 * Cp_water + 0.2 * Cp_glycol
                .Range("D5") = 1000
            ElseIf WorksheetFunction.Min(weatherWS.Range("F4", "F8764")) >= -19.3 Then
                .Range("E2") = Glycol30
                .Range("C5") = 0.7 * Cp_water + 0.3 * Cp_glycol
                .Range("D5") = 1000
            ElseIf WorksheetFunction.Min(weatherWS.Range("F4", "F8764")) >= -27.8 Then
                .Range("E2") = Glycol40
                .Range("C5") = 0.6 * Cp_water + 0.4 * Cp_glycol
                .Range("D5") = 1000
            ElseIf WorksheetFunction.Min(weatherWS.Range("F4", "F8764")) >= -30 Then
                .Range("E2") = Glycol42
                .Range("C5") = 0.58 * Cp_water + 0.42 * Cp_glycol
                .Range("D5") = 1000
            ElseIf WorksheetFunction.Min(weatherWS.Range("F4", "F8764")) >= -45 Then
                .Range("E2") = Glycol50
                .Range("C5") = 0.5 * Cp_water + 0.5 * Cp_glycol
                .Range("D5") = 1000
            End If
        Else
            .Range("E2") = Therminol66
        End If
    End With
End Sub
