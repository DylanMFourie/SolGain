VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTFPropertiesClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rho As Double    'density [kg/m^3]
Public c_p As Double    'Heat capacity [J/(kg*K)]
Public nu As Double     'Viscosity

Sub constructor(rho_kgPerM3 As Double, c_p_JPerkgK As Double)
    rho = rho_kgPerM3
    c_p = c_p_JPerkgK
End Sub
