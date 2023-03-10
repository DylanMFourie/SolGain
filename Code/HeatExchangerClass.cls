VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HeatExchangerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'HE characterisation declarations
Dim UA As Double

'HE outputs
Public T_hotOut As Double
Public T_coldOut As Double
Public T_hotIn As Double
Public T_coldIn As Double

Sub constructor(heatExchangerUA As Double)
    UA = heatExchangerUA
End Sub

Sub calcOutTemps(T_hotIn_K As Double, T_coldIn_K As Double, hotHTF As HTFPropertiesClass, coldHTF As HTFPropertiesClass, M_hot As Double, M_cold As Double)
    Dim C_hot As Double
    Dim C_cold As Double
    Dim C_max As Double
    Dim C_min As Double
    Dim NTU As Double   'Number of Transfer Units
    Dim effectiveness As Double
    Dim Q_exchanged As Double
    Dim helper1 As Double
    Dim helper2 As Double
    
    'Assign T_hotIn and T_coldIn
    T_hotIn = T_hotIn_K
    T_coldIn = T_coldIn_K
    
    'Calculate HTF heat transfer capacity (C_hot ans C_cold)
    C_hot = M_hot * hotHTF.c_p
    C_cold = M_cold * coldHTF.c_p
    
    'Determine C_max and C_min
    If C_hot > C_cold Then
        C_max = C_hot
        C_min = C_cold
    Else
        C_max = C_cold
        C_min = C_hot
    End If
    
    'Calculate Number of Transfer Units
    NTU = UA / C_max
    
    'Calculate HE effectiveness
    If (C_min / C_max) <> 1 Then
        helper1 = 1 - Exp(-NTU * (1 - C_min / C_max))
        helper2 = 1 - (C_min / C_max) * Exp(-NTU * (1 - C_min / C_max))
        effectiveness = helper1 / helper2
    Else
        effectiveness = NTU / (1 + NTU)
    End If
    
    'Calculate heat exchanged between HTFs
    Q_exchanged = effectiveness * C_min * (T_hotIn - T_coldIn)
    
    'Calculate outlet temperatures T_hotOut and T_coldOut
    T_hotOut = T_hotIn - Q_exchanged / C_hot
    T_coldOut = T_coldIn + Q_exchanged / C_cold
End Sub

Sub calcHotSideTemps(T_coldIn_K As Double, T_coldOut_K As Double, coldHTF As HTFPropertiesClass, hotHTF As HTFPropertiesClass, M_cold As Double, M_hot As Double)
    Dim C_hot As Double
    Dim C_cold As Double
    Dim C_max As Double
    Dim C_min As Double
    Dim NTU As Double   'Number of Transfer Units
    Dim effectiveness As Double
    Dim Q_exchanged As Double
    Dim helper1 As Double
    Dim helper2 As Double
    
    'Assign known temperatures
    T_coldIn = T_coldIn_K
    T_coldOut = T_coldOut_K
    
    'Calculate HTF heat transfer capacity (C_hot ans C_cold)
    C_hot = M_hot * hotHTF.c_p
    C_cold = M_cold * hotHTF.c_p
    
    'Determine C_max and C_min
    If C_hot > C_cold Then
        C_max = C_hot
        C_min = C_cold
    Else
        C_max = C_cold
        C_min = C_hot
    End If
    
    'Calculate Number of Transfer Units
    NTU = UA / C_max
    
    'Calculate HE effectiveness
    If (C_min / C_max) <> 1 Then
        helper1 = 1 - Exp(-NTU * (1 - C_min / C_max))
        helper2 = 1 - (C_min / C_max) * Exp(-NTU * (1 - C_min / C_max))
        effectiveness = helper1 / helper2
    Else
        effectiveness = NTU / (1 + NTU)
    End If
    
    'Calculate heat exchanged between HTFs
    Q_exchanged = C_cold * (T_coldOut - T_coldIn)
    
    'Calculate outlet temperatures T_hotIn and T_hotOut
    T_hotIn = T_coldIn + Q_exchanged / (effectiveness * C_min)
    T_hotOut = T_hotIn - Q_exchanged / C_hot
End Sub

