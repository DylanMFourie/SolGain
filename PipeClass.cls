VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PipeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Pipe characterization vars
Dim U As Double 'Heat loss coefficient [W/(K*m^2)]
Dim d As Double 'Pipe diameter [m]
Dim L As Double 'Pipe Length [m]

'Internal vars
Dim a As Double 'Pipe surface area [m^2]

Sub constructor(HeatLossCoeff As Double, Diameter As Double, Length As Double)
    U = HeatLossCoeff
    d = Diameter
    L = Length
    
    a = L * Pi * d ^ 2 / 4
End Sub

Function calcOutletTemp_K(InletTemp_K As Double, AmbientTemp_K As Double, HTF As HTFPropertiesClass, MassFlowRate As Double)
    Dim deltaT As Double
    Dim OutletTemp
    
    deltaT = (U * a * (InletTemp_K - AmbientTemp_K)) / (MassFlowRate * HTF.c_p)
    OutletTemp = InletTemp_K - deltaT
    calcOutletTemp_K = OutletTemp
End Function

Function calcInletTemp_K(desiredOutletTemp_K As Double, AmbientTemp_K As Double, HTF As HTFPropertiesClass, MassFlowRate As Double)
    Dim deltaT As Double
    Dim InletTemp As Double
    Dim InletTemp_prev As Double
    
    'Inititally sssume inlet temp = outlet temp
    InletTemp = desiredOutletTemp_K
    InletTemp_prev = 0  'Arbritrary number
    
    Do While Abs(InletTemp - InletTemp_prev) > iterx
        InletTemp_prev = InletTemp
        deltaT = (U * a * (InletTemp - AmbientTemp_K)) / (MassFlowRate * HTF.c_p)
        InletTemp = desiredOutletTemp_K + deltaT
    Loop
    
    calcInletTemp_K = InletTemp
End Function
