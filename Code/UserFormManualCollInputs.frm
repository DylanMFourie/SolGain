VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormManualCollInputs 
   Caption         =   "Collector Efficiency Paramaters"
   ClientHeight    =   9096.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10908
   OleObjectBlob   =   "UserFormManualCollInputs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormManualCollInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim collInWS As Worksheet

Private Sub ComboBoxCollectorType_Change()
    If ComboBoxCollectorType = ParabolicTrough Then
        TextBoxCollLength.Enabled = True
        TextBoxCollLength.BackColor = White
        TextBoxCollWidth.Enabled = True
        TextBoxCollWidth.BackColor = White
        TextBoxFocalLength.Enabled = True
        TextBoxFocalLength.BackColor = White
        
        ComboBoxNoReferencesTrans.Enabled = False
        ComboBoxNoReferencesTrans.BackColor = Grey
        ComboBoxNoReferencesTrans.Value = ""
    Else
        TextBoxFocalLength.Enabled = False
        TextBoxFocalLength.BackColor = Grey
        
        ComboBoxNoReferencesTrans.Enabled = True
        ComboBoxNoReferencesTrans.BackColor = White
    End If
End Sub

Private Sub ComboBoxNoReferencesTrans_Change()
    Dim i As Integer
    
    For i = 1 To 10
        If IsNumeric(ComboBoxNoReferencesTrans) = False Or i > ComboBoxNoReferencesTrans.Value Then
            UserFormManualCollInputs.Controls("TransThetaTextBox" & i).Enabled = False
            UserFormManualCollInputs.Controls("TransThetaTextBox" & i).BackColor = Grey
            UserFormManualCollInputs.Controls("TransThetaTextBox" & i).Value = ""
            UserFormManualCollInputs.Controls("TransKTextBox" & i).Enabled = False
            UserFormManualCollInputs.Controls("TransKTextBox" & i).BackColor = Grey
            UserFormManualCollInputs.Controls("TransKTextBox" & i).Value = ""
        Else
            UserFormManualCollInputs.Controls("TransThetaTextBox" & i).Enabled = True
            UserFormManualCollInputs.Controls("TransThetaTextBox" & i).BackColor = White
            UserFormManualCollInputs.Controls("TransKTextBox" & i).Enabled = True
            UserFormManualCollInputs.Controls("TransKTextBox" & i).BackColor = White
        End If
    Next i
    
End Sub


Private Sub ComboBoxNoReferencesLong_Change()
    Dim i As Integer
    
    For i = 1 To 10
        If IsNumeric(ComboBoxNoReferencesLong) = False Or i > ComboBoxNoReferencesLong.Value Then
            UserFormManualCollInputs.Controls("LongThetaTextBox" & i).Enabled = False
            UserFormManualCollInputs.Controls("LongThetaTextBox" & i).BackColor = Grey
            UserFormManualCollInputs.Controls("LongThetaTextBox" & i).Value = ""
            UserFormManualCollInputs.Controls("LongKTextBox" & i).Enabled = False
            UserFormManualCollInputs.Controls("LongKTextBox" & i).BackColor = Grey
            UserFormManualCollInputs.Controls("LongKTextBox" & i).Value = ""
        Else
            UserFormManualCollInputs.Controls("LongThetaTextBox" & i).Enabled = True
            UserFormManualCollInputs.Controls("LongThetaTextBox" & i).BackColor = White
            UserFormManualCollInputs.Controls("LongKTextBox" & i).Enabled = True
            UserFormManualCollInputs.Controls("LongKTextBox" & i).BackColor = White
        End If
    Next i
End Sub


Private Sub CommandButtonNext_Click()
    Dim i As Integer
    
    collInWS.Range("K2", "AA11").Clear
    
    If ComboBoxCollectorType = "" Then
        MsgBox "Please select collector type"
        Exit Sub
    ElseIf IsNumeric(TextBoxApArea) = False Then
        MsgBox "Please enter valid aperture area"
        Exit Sub
    ElseIf TextBoxCollLength.Enabled = True And IsNumeric(TextBoxCollLength) = False Then
        MsgBox "Please enter valid collector length"
        Exit Sub
    ElseIf TextBoxCollWidth.Enabled = True And IsNumeric(TextBoxCollWidth) = False Then
        MsgBox "Please enter valid collector width"
        Exit Sub
    ElseIf TextBoxFocalLength.Enabled = True And IsNumeric(TextBoxFocalLength) = False Then
        MsgBox "Please enter valid focal length"
        Exit Sub
    ElseIf IsNumeric(TextBox_n_0) = False Or TextBox_n_0 > 100# Then
        MsgBox "Please enter valid Optical Efficiency (n_0)"
        Exit Sub
    ElseIf IsNumeric(TextBox_c_1) = False Then
        MsgBox "Please enter valid 1st Order Heat Loss Coefficient (c_1)"
        Exit Sub
    ElseIf IsNumeric(TextBox_c_2) = False Then
        MsgBox "Please enter valid 2nd Order Heat Loss Coefficient (c_2)"
        Exit Sub
    ElseIf IsNumeric(TextBox_c_eff) = False Or TextBox_c_eff > 100 Then
        MsgBox "Please enter valid Collector Heat capacity (c_eff)"
        Exit Sub
    ElseIf IsNumeric(TextBox_K_d) = False Or TextBox_K_d > 1 Then
        MsgBox "Please enter valid Diffuse Incidence Angle Modifier (K_d)"
        Exit Sub
    ElseIf ComboBoxNoReferencesTrans.Enabled = True And IsNumeric(ComboBoxNoReferencesTrans) = False Then
        MsgBox "Please select number of transversal Incidence Angle Modifier references"
        Exit Sub
    ElseIf ComboBoxNoReferencesLong.Enabled = True And IsNumeric(ComboBoxNoReferencesLong) = False Then
        MsgBox "Please select number of longitudinal Incidence Angle Modifier references"
        Exit Sub
    End If
    
    If ComboBoxNoReferencesTrans.Enabled = True Then
        For i = 1 To ComboBoxNoReferencesTrans.Value
            If IsNumeric(UserFormManualCollInputs.Controls("TransThetaTextBox" & i)) = False Or IsNumeric(UserFormManualCollInputs.Controls("TransKTextBox" & i)) = False Then
                MsgBox "Please enter valid transversal reference IAMs"
                Exit Sub
            End If
            If UserFormManualCollInputs.Controls("TransThetaTextBox" & i) > 90 Or UserFormManualCollInputs.Controls("TransThetaTextBox" & i) < 0 Then
                MsgBox "Transversal reference angles must be between 0 and 90" & Chr(176)
                Exit Sub
            End If
            If UserFormManualCollInputs.Controls("TransKTextBox" & i) > 2 Or UserFormManualCollInputs.Controls("TransKTextBox" & i) < 0 Then
                MsgBox "Transversal reference IAMs must be between 0 and 2"
            End If
        Next i
    End If
    
    If ComboBoxNoReferencesLong.Enabled = True Then
        For i = 1 To ComboBoxNoReferencesLong.Value
            If IsNumeric(UserFormManualCollInputs.Controls("LongThetaTextBox" & i)) = False Or IsNumeric(UserFormManualCollInputs.Controls("LongKTextBox" & i)) = False Then
                MsgBox "Please enter longitudinal reference IAMs"
                Exit Sub
            End If
            If UserFormManualCollInputs.Controls("LongThetaTextBox" & i) > 90 Or UserFormManualCollInputs.Controls("LongThetaTextBox" & i) < 0 Then
                MsgBox "Longitudinal reference angles must be between 0 and 90" & Chr(176)
                Exit Sub
            End If
            If UserFormManualCollInputs.Controls("LongKTextBox" & i) > 2 Or UserFormManualCollInputs.Controls("LongKTextBox" & i) < 0 Then
                MsgBox "Longitudinal references IAMs must be between 0 and 2"
            End If
        Next i
    End If
    
    
    If ComboBoxCollectorType = ParabolicTrough Then
        collInWS.Range("K2") = True
    Else
        collInWS.Range("K2") = False
    End If
    collInWS.Range("L2") = ComboBoxCollectorType
    
    collInWS.Range("M2") = Val(TextBoxApArea)
    collInWS.Range("N2") = Val(TextBoxCollLength)
    collInWS.Range("O2") = Val(TextBoxCollWidth)
    collInWS.Range("P2") = Val(TextBoxFocalLength)
    
    collInWS.Range("Q2") = Val(TextBox_n_0)
    collInWS.Range("R2") = Val(TextBox_c_1)
    collInWS.Range("S2") = Val(TextBox_c_2)
    
    collInWS.Range("T2") = Val(TextBox_c_eff) / 1000
    collInWS.Range("U2") = Val(TextBox_K_d)
    
    collInWS.Range("V2") = Val(ComboBoxNoReferencesTrans)
    collInWS.Range("Y2") = Val(ComboBoxNoReferencesLong)
    
    If ComboBoxNoReferencesTrans.Value <> "" Then
        For i = 1 To ComboBoxNoReferencesTrans.Value
            collInWS.Range("W" & i + 1) = Val(UserFormManualCollInputs.Controls("TransThetaTextBox" & i))
            collInWS.Range("X" & i + 1) = Val(UserFormManualCollInputs.Controls("TransKTextBox" & i))
        Next i
    End If
    
    If ComboBoxNoReferencesLong.Value <> "" Then
        For i = 1 To ComboBoxNoReferencesLong.Value
            collInWS.Range("Z" & i + 1) = Val(UserFormManualCollInputs.Controls("LongThetaTextBox" & i))
            collInWS.Range("AA" & i + 1) = Val(UserFormManualCollInputs.Controls("LongKTextBox" & i))
        Next i
    End If
    
    'Assign HTF in collector loop and assign properties
    Call UserFormFirstCollector.assignColHTF
    
    UserFormManualCollInputs.Hide
    
    If ComboBoxCollectorType.Value = ParabolicTrough Then
        UserFormTiltOrientationPTC.Show
    ElseIf ComboBoxCollectorType.Value = ETC Then
        UserFormTiltOrientationETC.Show
    ElseIf ComboBoxCollectorType.Value = FPC Then
        UserFormTiltOrientationFPC.Show
    End If
End Sub

Private Sub CommandButtonPrev_Click()
    Unload UserFormManualCollInputs
    UserFormFirstCollector.Show
End Sub



Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Set collInWS = ThisWorkbook.Sheets("Collector Inputs")
    
    For i = 1 To 10
        ComboBoxNoReferencesTrans.AddItem i
        ComboBoxNoReferencesLong.AddItem i
        UserFormManualCollInputs.Controls("LongThetaTextBox" & i).Enabled = False
        UserFormManualCollInputs.Controls("LongThetaTextBox" & i).BackColor = Grey
        UserFormManualCollInputs.Controls("LongKTextBox" & i).Enabled = False
        UserFormManualCollInputs.Controls("LongKTextBox" & i).BackColor = Grey
        UserFormManualCollInputs.Controls("TransThetaTextBox" & i).Enabled = False
        UserFormManualCollInputs.Controls("TransThetaTextBox" & i).BackColor = Grey
        UserFormManualCollInputs.Controls("TransKTextBox" & i).Enabled = False
        UserFormManualCollInputs.Controls("TransKTextBox" & i).BackColor = Grey
    Next i
    
    With ComboBoxCollectorType
        .AddItem FPC
        .AddItem ETC
        .AddItem ParabolicTrough
    End With
    
    For i = 1 To 20
        UserFormManualCollInputs.Controls("LabelDeg" & i) = "[" & Chr(176) & "]"
    Next i
End Sub


