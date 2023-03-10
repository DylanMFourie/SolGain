VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemperaturesClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'System Temperatures (**temps in degree K)
Public collectorOut_K As Double
Public collectorIn_K As Double
Public HEToCollector_K As Double
Public HEFromCollector_K As Double
Public HEToTank_K As Double
Public HEFromTank_K As Double
Public auxIn_K As Double
Public auxNoSolar_K As Double
Public processIn_K As Double
Public processOut_K As Double

'Environment Temperatures (**temps in degree K)
Public ambient_K As Double
Public ambientTank_K As Double

'Error
Public error As Double

Sub constructor(T_processInDesired As Double, T_processOut As Double, T_ambient_DegC As Double, T_ambientTank_DegC As Double)
    
    'Convert input Temps in Degree Celsius to Kelvin
    collectorOut_K = T_ambient_DegC + kelvin0
    collectorIn_K = T_ambient_DegC + kelvin0
    HEToCollector_K = T_ambient_DegC + kelvin0
    HEFromCollector_K = T_ambient_DegC + kelvin0
    HEToTank_K = T_ambient_DegC + kelvin0
    HEFromTank_K = T_ambient_DegC + kelvin0
    auxIn_K = T_ambientTank_DegC + kelvin0
    auxNoSolar_K = T_ambientTank_DegC + kelvin0
    processIn_K = T_processInDesired + kelvin0
    processOut_K = T_processOut + kelvin0
    
    ambient_K = T_ambient_DegC + kelvin0
    ambientTank_K = T_ambientTank_DegC + kelvin0
End Sub

Sub updateAmbientTemp(T_ambient_DegC As Double)
    ambient_K = T_ambient_DegC + kelvin0
End Sub

