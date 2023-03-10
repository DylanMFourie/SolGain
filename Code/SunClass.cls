VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SunClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**** All angles are in degrees unless otherwise stated.

'Sun characterization declarations
Dim albedo As Double
Dim latitude As Double
Dim longitudeCorrection As Double

'Sun hourly outputs declarations
Public G_bt As Double
Public G_st As Double
Public G_rt As Double
Public incidenceAngle As Double
Public incidenceAngle_longitudinal As Double
Public incidenceAngle_transverse As Double
Public azimuth As Double
Public zenith As Double
Public d_local As Single
Public h_local As Single

Sub constructor(grd_reflectance As Double, lat As Double, longitude As Double, TimeZoneLongitude As Double)
    albedo = grd_reflectance
    latitude = lat
    
    longitudeCorrection = (longitude - TimeZoneLongitude) / 15
End Sub

Sub calcIrradianceComps(G_d As Double, G_bn As Double, collData As CollectorClass)
    Dim helper As Double
    Dim elevationAngle As Double
    
    elevationAngle = 90 - zenith
    
    G_bt = 0
    G_st = 0
    G_rt = 0
    
    'If the sun is up then
    If zenith < 90 Then
        If collData.concentrating = False Then
            helper = Cos(deg2rad * zenith) * Cos(deg2rad * collData.tilt) + Sin(deg2rad * zenith) * Sin(deg2rad * collData.tilt) * Cos(deg2rad * (azimuth - collData.azimuth_coll))
            incidenceAngle = WorksheetFunction.Acos(helper) / deg2rad
            
            If collData.cType = ETC And collData.tubesVertical = False Then
                incidenceAngle_transverse = Atn(Tan(deg2rad * incidenceAngle) * Cos(deg2rad * (azimuth - collData.azimuth_coll))) / deg2rad
                incidenceAngle_longitudinal = Atn(Tan(deg2rad * incidenceAngle) * Sin(deg2rad * (azimuth - collData.azimuth_coll))) / deg2rad
            Else
                incidenceAngle_longitudinal = Atn(Tan(deg2rad * incidenceAngle) * Cos(deg2rad * (azimuth - collData.azimuth_coll))) / deg2rad
                incidenceAngle_transverse = Atn(Tan(deg2rad * incidenceAngle) * Sin(deg2rad * (azimuth - collData.azimuth_coll))) / deg2rad
            End If
        Else
            helper = Cos(deg2rad * (elevationAngle - collData.tilt)) - Cos(deg2rad * collData.tilt) * Cos(deg2rad * elevationAngle) * (1 - Cos(deg2rad * (azimuth - collData.azimuth_coll)))
            incidenceAngle = WorksheetFunction.Acos(Sqr(1 - helper ^ 2)) / deg2rad
            incidenceAngle_longitudinal = incidenceAngle
        End If
    Else
        incidenceAngle = 90
    End If
    
    If (G_d + G_bn) > 0.1 Then
        If incidenceAngle < 90 Then
            G_bt = G_bn * Cos(deg2rad * incidenceAngle)
        End If
        G_st = G_d * 0.5 * (1 + Cos(deg2rad * collData.tilt))
        G_rt = (G_bn + G_d) * (albedo / 2) * (1 - Cos(deg2rad * collData.tilt))
    End If
End Sub

'calcSunPos() is calculated in accordance with Technical Manual for the SAM Physical Trough Model
Sub calcSunPos(IterationNumber As Integer)
    Dim B As Double
    Dim EOT As Double
    Dim solarTime As Double
    Dim hourAngle As Double
    Dim declination As Double
    Dim EQ As Double
    Dim helper1 As Double
    Dim helper2 As Double
    
    'Determine day and hour
    d_local = Int((IterationNumber - 1) / 24) + 1
    h_local = (IterationNumber + 24) - d_local * 24 - 0.5
    
    'Determine day of equinox (EQ)
    If Sgn(latitude) = 1 Or latitude = 0 Then
        EQ = 284
    ElseIf Sgn(latitude) = -1 Then
        EQ = 101
    End If
    
    'Calculate Solar Time
    B = (360# / 365#) * (d_local - 1)
    EOT = 229.2 * (0.000075 + 0.001868 * Cos(deg2rad * B) - 0.032077 * Sin(deg2rad * B) - 0.014615 * Cos(2 * deg2rad * B) - 0.04089 * Sin(2 * deg2rad * B))
    solarTime = h_local + longitudeCorrection + EOT / 60#
    
    'Calculate Zenith Angle
    hourAngle = (solarTime - 12) * 15
    
    declination = 23.45 * Sin(deg2rad * 360# * ((EQ + d_local) / 365))
    'declination = (180 / Pi) * (0.006918 - 0.399912 * Cos(deg2rad * B) + 0.070257 * Sin(deg2rad * B) - 0.006758 * Cos(2 * deg2rad * B) + 0.000907 * Sin(2 * deg2rad * B) - 0.002697 * Cos(3 * deg2rad * B) + 0.00148 * Sin(3 * deg2rad * B))
    
    helper1 = Cos(deg2rad * declination) * Cos(deg2rad * hourAngle) * Cos(deg2rad * Abs(latitude)) + Sin(deg2rad * declination) * Sin(deg2rad * Abs(latitude))
    zenith = 90 - WorksheetFunction.Asin(helper1) / deg2rad
    
    'Calculate Azimuth Angle **** Measured clockwise from due north
    azimuth = 0
    
    If zenith < 90 Then 'Only calculate Azimuth if sun is up
        helper1 = (Cos(deg2rad * zenith) * Sin(deg2rad * Abs(latitude)) - Sin(deg2rad * declination)) / (Sin(deg2rad * zenith) * Cos(deg2rad * Abs(latitude)))
        azimuth = Sgn(hourAngle) * WorksheetFunction.Acos(helper1) / deg2rad
    End If
End Sub
