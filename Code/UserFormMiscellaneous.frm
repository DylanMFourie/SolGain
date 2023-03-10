VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMiscellaneous 
   Caption         =   "Miscellaneous Inputs"
   ClientHeight    =   7740
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13788
   OleObjectBlob   =   "UserFormMiscellaneous.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormMiscellaneous"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miscWS As Worksheet
Dim collWS As Worksheet

Private Sub ComboBoxHeatStorage_Change()
    If ComboBoxHeatStorage.Value = "Yes" Then
        ImageTank.Visible = True
        ImageNoTank.Visible = False
    ElseIf ComboBoxHeatStorage.Value = "No" Then
        ImageTank.Visible = False
        ImageNoTank.Visible = True
        
        With TextBoxStoreVolume
            .Enabled = False
            .BackColor = Grey
            .Value = 0
        End With
        
        With TextBoxStorageHLCoeff
            .Enabled = False
            .BackColor = Grey
            .Value = 0
        End With
    End If
End Sub

Private Sub CommandButtonNext_Click()
    
    If ComboBoxHeatStorage = "" Then
        MsgBox "Please select whether there is Heat Storage or Not"
        Exit Sub
    ElseIf IsNumeric(TextBoxStoreVolume) = False Then
        MsgBox "Please enter valid Heat Storage Volume"
        Exit Sub
    ElseIf IsNumeric(TextBoxStorageHLCoeff) = False Then
        MsgBox "Please enter valid Heat Storage Heat Loss Coefficient"
        Exit Sub
    ElseIf IsNumeric(TextBoxHeatExchangerUA) = False Then
        MsgBox "Please enter valid Heat Exchanger UA"
        Exit Sub
    ElseIf IsNumeric(TextBoxPipeHLCoeff) = False Then
        MsgBox "Please enter valid Pipe Heat Loss Coefficient"
        Exit Sub
    ElseIf IsNumeric(TextBoxPipeDiameter) = False Then
        MsgBox "Please enter valid Pipe Diameter"
        Exit Sub
    ElseIf IsNumeric(TextBoxDistBetweenCollAndTank) = False Then
        MsgBox "Please enter valid Distance between Collector and Heat Exchanger"
        Exit Sub
    End If
    
    miscWS.Range("B2") = ComboBoxHeatStorage
    miscWS.Range("C2") = Val(TextBoxStoreVolume)
    miscWS.Range("D2") = Val(TextBoxStorageHLCoeff)
    miscWS.Range("E2") = Val(TextBoxHeatExchangerUA)
    miscWS.Range("F2") = Val(TextBoxPipeHLCoeff)
    miscWS.Range("G2") = Val(TextBoxPipeDiameter)
    miscWS.Range("H2") = Val(TextBoxDistBetweenCollAndTank)
    
    UserFormMiscellaneous.Hide
    
    Call Simulation
End Sub

Private Sub CommandButtonPrev_Click()
    UserFormMiscellaneous.Hide
    UserFormYearlyDemand.Show
End Sub




Private Sub UserForm_Initialize()
    Dim M_proMax As Double
    Dim rho_pro As Double
    Dim area_HE As Double
    Dim U_HE As Double
    Dim area_collField As Double
    Dim M_coll_kgperh As Double
    Dim rho_coll As Double
    Dim V_coll_lperh As Double
    
    Set miscWS = ThisWorkbook.Sheets("Misc Inputs")
    Set collWS = ThisWorkbook.Sheets("Collector Inputs")
    
    M_proMax = ThisWorkbook.Sheets("Process Inputs").Range("D2")
    rho_pro = ThisWorkbook.Sheets("Process Inputs").Range("E2")
    
    If collWS.Range("A5") = True Then
        area_collField = collWS.Range("F2") * collWS.Range("G2") * collWS.Range("M2")
    Else
        area_collField = collWS.Range("B5")
    End If
    
    ComboBoxHeatStorage.AddItem "Yes"
    ComboBoxHeatStorage.AddItem "No"
    ComboBoxHeatStorage.Value = "Yes"
    
    ImageNoTank.Visible = False
    
    'Assume Storage Volume
    TextBoxStoreVolume.Value = 1.2 * 24 * (M_proMax / rho_pro)
    
    'Assume Storage HL coeff
    TextBoxStorageHLCoeff.Value = 0.3
    
    'Assume HE heat transfer coeff.
    area_HE = 0.2 * area_collField
    U_HE = 500
    TextBoxHeatExchangerUA.Value = U_HE * area_HE
    
    'Assume pipe HL coeff.
    TextBoxPipeHLCoeff.Value = 0.8
    
    'Assume Pipe Diameter
    M_coll_kgperh = 18 * area_collField
    rho_coll = collWS.Range("D5")
    V_coll_lperh = 1000 * (M_coll_kgperh / rho_coll)
    TextBoxPipeDiameter.Value = Sqr(0.35 * V_coll_lperh)

    'Assume distance between collector and Tank
    TextBoxDistBetweenCollAndTank.Value = 10
End Sub
