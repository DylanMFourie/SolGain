VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MassFlowClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Outputs
Public M_collector As Double
Public M_coll_prev As Double
Public M_process As Double

'System Characterization
Dim M_proMax As Double
Public heatingStrategy As String
Public T_processTarget_K As Double
Dim T_processOut_K As Double

'In object variables
Dim Q_max As Double

Public Sub constructor(M_pro_KGperH As Double, c_p_process As Double, T_processIn_C As Double, T_processOut_C As Double, heatingStrat As String)
    M_proMax = M_pro_KGperH / (60# * 60)
    
    Q_max = M_proMax * c_p_process * (T_processIn_C - T_processOut_C)
    
    T_processTarget_K = T_processIn_C + kelvin0
    T_processOut_K = T_processOut_C + kelvin0
    
    M_process = M_proMax
    M_collector = M_proMax
    
    heatingStrategy = heatingStrat
End Sub

'Sub Constructor(hourlyPercentages As Range, dailyPercentages As Range, weeklyPercentages As Range, maxHeatDemand As Double)
''    ReDim hourlyPercentages(1 To 24) As Double
''    ReDim dailyPercentages(1 To 7) As Double
''    ReDim weeklyPercentages(1 To 53) As Double
'
'    hourlyDemands = hourlyPercentages.Value
'    dailyDemands = dailyPercentages.Value
'    weeklyDemands = weeklyPercentages.Value
'
'    maxHeatConsumption = maxHeatDemand
'End Sub
'

Function calcProcessFlowRate(hourlyDemand As Double, dailyDemand As Double, weeklyDemand As Double)
    Dim processMassFlow As Double
    
    processMassFlow = (hourlyDemand * dailyDemand * weeklyDemand) * M_proMax
    
    calcProcessFlowRate = processMassFlow
End Function
