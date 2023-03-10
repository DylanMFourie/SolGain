VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectorClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'System constants
Private Const minSpecificFlowRate As Double = 0.001 'Minimum specific collector flow rate that does not result in numerical problems

'Collector characterization declarations
Public concentrating As Boolean
Public cType As String
Public azimuth_coll As Double
Public tilt As Double
Public tubesVertical As Boolean
Public n_parallel_modules As Integer  'Number of modules in parallel
Public n_series_coll As Integer    'Number of collectors in series
Public apertureArea_singleColl As Double  'Aperture area [m^2]
Public grossApertureArea_field As Double
Public manualInputs As Boolean
Dim length_coll As Double
Dim width_coll As Double
Dim focalLength As Double
Dim n_rows As Double
Dim n_columns As Double
Dim distanceBetweenRows As Double

'Collector efficiency parameters
Dim n_0 As Double   'collecotr efficiency
Dim c_1 As Double   '[W/(m²*K)] C1
Dim c_2 As Double   '[W/(m²*K²)] C2
Dim theta_ref_l() As Double   '[°]refernce angle of collector
Dim K_ref_l() As Double   'reference IAM
Dim theta_ref_t() As Double   '[°]refernce angle of collector
Dim K_ref_t() As Double   'reference IAM
Dim numberIAMrefs_l As Integer
Dim numberIAMrefs_t As Integer
Dim c_eff As Double '[J/(m²*K)] effective thermal capacity
Dim K_d As Double   'IAM for diffuse irradiation

'In object parameters
Dim T_f_prev() As Double
Dim K_b As Double
Dim K_s As Double
Dim K_r As Double
Dim collWS As Worksheet

Sub constructor(concentratingTrueFalse As Boolean, collectorType As String, n_collSeries As Integer, n_modulesParallel As Integer, T_ambient_K As Double)
    Dim i As Integer
    
    Set collWS = ThisWorkbook.Sheets("Collector Inputs")
    
    concentrating = concentratingTrueFalse
    cType = collectorType
    
    With collWS
        tilt = .Range("A2")
        azimuth_coll = .Range("B2")
        'tubesVertical = True/False
        n_series_coll = n_collSeries
        n_parallel_modules = n_modulesParallel
        n_rows = .Range("H2")
        n_columns = .Range("I2")
        distanceBetweenRows = .Range("J2")
        
        apertureArea_singleColl = .Range("M2")
        length_coll = .Range("N2")
        width_coll = .Range("O2")
        grossApertureArea_field = n_parallel_modules * n_series_coll * (apertureArea_singleColl)
        focalLength = .Range("P2")
        
        n_0 = .Range("Q2")  'collecotr efficiency
        c_1 = .Range("R2")  '[W/(m²*K)] C1
        c_2 = .Range("S2")  '[W/(m²*K²)] C2
        c_eff = .Range("T2")    '[J/(m²*K)] effective thermal capacity
        K_d = .Range("U2")  'IAM for diffuse irradiation
        
        manualInputs = .Range("A5")
    End With
    
'    n_0 = 0.811  'collecotr efficiency
'    c_1 = 2.71   '[W/(m²*K)] C1
'    c_2 = 0.01   '[W/(m²*K²)] C2
'        theta_ref_l = 50# '[°]refernce angle of collector
'        K_ref_l = 0.96  'reference IAM for Flat Plate Collector
'        theta_ref_t = 50# '[°]refernce angle of collector
'        K_ref_t = 0.96 'reference IAM for Flat Plate Collector
'    c_eff = 7050# '[J/(m²*K)] effective thermal capacity
'    K_d = 0.912  'IAM for diffuse irradiation`
    
    Call inputIAMReferrences
    
    ReDim T_f_prev(1 To n_series_coll)
    
    For i = 1 To n_series_coll
        T_f_prev(i) = T_ambient_K
    Next i
End Sub

Sub inputIAMReferrences()
    Dim i As Integer
    
    numberIAMrefs_l = collWS.Range("Y2")
    numberIAMrefs_t = collWS.Range("V2")
    
    ReDim K_ref_l(0 To numberIAMrefs_l + 1)
    ReDim theta_ref_l(0 To numberIAMrefs_l + 1)
    ReDim K_ref_t(0 To numberIAMrefs_t + 1)
    ReDim theta_ref_t(0 To numberIAMrefs_t + 1)
    
    K_ref_l(0) = 1
    K_ref_t(0) = 1
    theta_ref_l(0) = 0
    theta_ref_t(0) = 0
    
    K_ref_l(numberIAMrefs_l + 1) = 0
    K_ref_t(numberIAMrefs_t + 1) = 0
    theta_ref_l(numberIAMrefs_l + 1) = 90
    theta_ref_t(numberIAMrefs_t + 1) = 90
    
    For i = 1 To numberIAMrefs_l
        theta_ref_l(i) = collWS.Range("Z" & i + 1)
        K_ref_l(i) = collWS.Range("AA" & i + 1)
    Next i
    
    For i = 1 To numberIAMrefs_t
        theta_ref_t(i) = collWS.Range("W" & i + 1)
        K_ref_t(i) = collWS.Range("X" & i + 1)
    Next i
End Sub

Sub calcTilt(sunData As SunClass)
    tilt = tilt
End Sub


'This function uses the bisection method to calculate the flow rate for a desired output temp
Function calcFlowRate(sun As SunClass, T_In_K As Double, T_outDesired_K As Double, T_ambient_K As Double, HTF As HTFPropertiesClass, approximateMassFlow As Double) As Double
    'f(M) = T_OutDesired - T(M)
    
    Dim f_up As Double
    Dim f_mid As Double
    Dim f_low As Double
    Dim M_up As Double
    Dim M_mid As Double
    Dim M_low As Double
    
    If approximateMassFlow / (apertureArea_singleColl * n_parallel_modules) = minSpecificFlowRate Then
        M_mid = approximateMassFlow
    Else
        M_low = minSpecificFlowRate * (apertureArea_singleColl * n_parallel_modules)
        M_up = 10 * approximateMassFlow
        M_mid = (M_low + M_up) / 2
        
        f_low = flowError(sun, T_In_K, T_outDesired_K, T_ambient_K, HTF, M_low)
        f_up = flowError(sun, T_In_K, T_outDesired_K, T_ambient_K, HTF, M_up)
        f_mid = flowError(sun, T_In_K, T_outDesired_K, T_ambient_K, HTF, M_mid)
        
        'Loop until T_Out_K_Desired is found
        Do While Abs(f_mid) > iterx And (M_mid - minSpecificFlowRate * apertureArea_singleColl * n_parallel_modules) > iterx 'And (Abs(M_up - M_low) / 2) > iterx
            
            If Sgn(f_mid) = Sgn(f_up) Then
                M_up = M_mid
            ElseIf Sgn(f_mid) = Sgn(f_low) Then
                M_low = M_mid
            End If
            
            M_mid = (M_up + M_low) / 2
            
            f_low = flowError(sun, T_In_K, T_outDesired_K, T_ambient_K, HTF, M_low)
            f_up = flowError(sun, T_In_K, T_outDesired_K, T_ambient_K, HTF, M_up)
            f_mid = flowError(sun, T_In_K, T_outDesired_K, T_ambient_K, HTF, M_mid)
        Loop
    End If
    
    'Surrender to Him, Abide in Him, Rest in Him
    
    'Output of function
    calcFlowRate = M_mid
End Function

Private Function flowError(sun As SunClass, T_In_K As Double, T_outDesired_K As Double, T_ambient_K As Double, HTF As HTFPropertiesClass, MassFlowRate As Double) As Double
    Dim T_out_i As Double
    Dim T_in_i As Double
    Dim specificMassFlow As Double '[kg/(s*m^2)]
    Dim i As Integer
    Dim n As Integer
    
    'Calculate specific mass flow rate [kg/(s*m^2)]
    specificMassFlow = MassFlowRate / (apertureArea_singleColl * n_parallel_modules)
    
    If specificMassFlow < 0.1 / (60# * 60#) Then 'specificMassFlow/n_parallel_modules [kg/(m^2*s) = (kg/(m^2*h))/(60mins*60secs)]
        n = 10
    Else
        n = n_series_coll
    End If
    
    T_out_i = T_In_K
    T_in_i = T_In_K
    
    'Calculate collector output temp
    For i = 1 To n  'Go through all collector in series up till the nth collector
        T_out_i = calcIndividualCollOut(i, sun, T_in_i, T_ambient_K, HTF, specificMassFlow)
        T_in_i = T_out_i    'Output temp. of collector becomes next collector in series' input temp.
    Next i
    
    'Output of function
    flowError = T_outDesired_K - T_out_i
End Function

Function calcOutputTemp(sun As SunClass, T_In_K As Double, T_ambient_K As Double, HTF As HTFPropertiesClass, MassFlowRate As Double) As Double
    Dim T_out_i As Double
    Dim T_in_i As Double
    Dim specificMassFlow As Double '[kg/(s*m^2)]
    Dim i As Integer
    Dim n As Integer
    
    'Calculate specific mass flow rate [kg/(s*m^2)]
    specificMassFlow = MassFlowRate / (apertureArea_singleColl * n_parallel_modules)
    
    If specificMassFlow < 0.1 / (60# * 60#) Then 'specificMassFlow/n_parallel_modules [kg/(m^2*s) = (kg/(m^2*h))/(60mins*60secs)]
        n = 10
    Else
        n = n_series_coll
    End If
    
    T_out_i = T_In_K
    T_in_i = T_In_K
    
    'Calculate collector output temp
    For i = 1 To n  'Go through all collector in series up till the nth collector
        T_out_i = calcIndividualCollOut(i, sun, T_in_i, T_ambient_K, HTF, specificMassFlow)
        
        T_f_prev(i) = (T_in_i + T_out_i) / 2#
        T_in_i = T_out_i    'Output temp. of collector becomes next collector in series' input temp.
    Next i
    
    'Output of function
    calcOutputTemp = T_out_i
End Function


Function approxFlowRate(sun As SunClass, T_collIn_K As Double, T_collOut_K As Double, T_ambient_K As Double, HTF As HTFPropertiesClass) As Double
    Dim T_f As Double
    Dim dT_f As Double
    Dim q_gain As Double
    Dim Q_coll As Double
    Dim approxEndLossEff As Double
    Dim approxM_coll As Double
    
    'Approximate average collector temp
    T_f = (T_collIn_K + T_collOut_K) / 2#
    
    'Calc dT_f using approx T_f and average T_f_prev
    dT_f = (T_f - WorksheetFunction.Average(T_f_prev)) / timeStep
    
    'Calc q_gain based on approx T_f
    If concentrating = True Then
        approxEndLossEff = calcAvgEndLossEfficiency(sun.incidenceAngle_longitudinal)
    Else
        approxEndLossEff = 1
    End If
    
    q_gain = n_0 * (approxEndLossEff * K_b * sun.G_bt + K_s * sun.G_st + K_r * sun.G_rt) - c_1 * (T_f - T_ambient_K) - c_2 * (T_f - T_ambient_K) ^ 2#
    
    'Calc Q_coll
    Q_coll = n_series_coll * (apertureArea_singleColl * (q_gain - c_eff * dT_f))
    
    'Approximate M_coll
    If Q_coll > 0# Then
        approxM_coll = (Q_coll / (HTF.c_p * (T_collOut_K - T_collIn_K))) * n_parallel_modules
    Else
        approxM_coll = minSpecificFlowRate * (apertureArea_singleColl * n_parallel_modules)
    End If
    
    'Determine whether in stagnation mode
    If approxM_coll / (apertureArea_singleColl * n_parallel_modules) < minSpecificFlowRate Then 'specificMassFlow/n_parallel_modules [kg/(m^2*s) = (kg/(m^2*h))/(60mins*60secs)]
        approxM_coll = minSpecificFlowRate * (apertureArea_singleColl * n_parallel_modules)
    End If
    
    approxFlowRate = approxM_coll
End Function


Private Function calcIndividualCollOut(i_collector As Integer, sun As SunClass, T_in_i As Double, T_ambient_K As Double, HTF As HTFPropertiesClass, specificMassFlow As Double) As Double
    Dim T_out As Double
    Dim T_out_prev As Double
    Dim T_f As Double
    Dim q_gain As Double
    Dim endLossEff As Double
    Dim helper1 As Double
    Dim helper2 As Double

    T_out_prev = 0#
    T_out = T_in_i

    'Calculate end loss efficiency
    If concentrating = True Then
        endLossEff = calcEndLossEfficiencyIndividualCollector(i_collector, sun.incidenceAngle_longitudinal)
    Else
        endLossEff = 1
    End If
    
    Do While Abs(T_out - T_out_prev) > iterx
        T_out_prev = T_out
        
        'Calculate average collector temp
        T_f = (T_in_i + T_out) / 2#
        
        'Calculate useful collector gain
        q_gain = n_0 * (endLossEff * K_b * sun.G_bt + K_s * sun.G_st + K_r * sun.G_rt) - c_1 * (T_f - T_ambient_K) - c_2 * (T_f - T_ambient_K) ^ 2#
        
        'Calculate collector outlet temperature
        helper1 = specificMassFlow * HTF.c_p * T_in_i - (c_eff / (2# * timeStep)) * T_in_i + q_gain + (c_eff / timeStep) * T_f_prev(i_collector)
        helper2 = specificMassFlow * HTF.c_p + c_eff / (2# * timeStep)
        T_out = helper1 / helper2
    Loop
    
    calcIndividualCollOut = T_out
End Function

'Private Function collError(i_coll As Integer, sun As SunClass, T_in_i As Double


Sub calcIAMs(sunData As SunClass)
    Dim K_b_l As Double
    Dim K_b_t As Double
    Dim theta_l As Double
    Dim theta_t As Double
    
    theta_l = sunData.incidenceAngle_longitudinal
    theta_t = sunData.incidenceAngle_transverse
    
    'Determine IAMs based on angle of sun and irradiance supply
    If (sunData.G_bt + sunData.G_rt + sunData.G_st) < 0.1 Then
        K_b = 0#
        K_s = 0#
        K_r = 0#
    ElseIf sunData.incidenceAngle >= 90# Then
        K_b = 0#
        If cType = ParabolicTrough Then
            K_s = 0#
            K_r = 0#
        Else
            K_s = K_d
            K_r = K_d
        End If
    Else
    
        'Calculate longitudinal IAM
        If numberIAMrefs_l = 1 Then
            K_b_l = Ambrosetti(theta_ref_l(1), K_ref_l(1), theta_l)
        Else
            K_b_l = Interpolation(numberIAMrefs_l, theta_ref_l, K_ref_l, theta_l)
        End If
        
        'Calculate transverse IAM
        If cType = ParabolicTrough Then
            K_b_t = 1
        Else
            If numberIAMrefs_t = 1 Then
                K_b_t = Ambrosetti(theta_ref_t(1), K_ref_t(1), theta_t)
            Else
                K_b_t = Interpolation(numberIAMrefs_t, theta_ref_t, K_ref_t, theta_t)
            End If
        End If
        
        'Calculate overall beam IAM
        K_b = K_b_t * K_b_l 'IAM dependent on Thetat and Thetal
        
        If cType = ParabolicTrough Then
            K_s = 0#
            K_r = 0#
        Else
            K_s = K_d
            K_r = K_d
        End If
    End If
End Sub

Private Function Ambrosetti(theta_ref As Double, K_ref As Double, theta As Double)
    Dim r As Double
    Dim k As Double
    
    r = Application.WorksheetFunction.Ln(Tan(deg2rad * (theta_ref / 2#))) / Application.WorksheetFunction.Ln(1# - K_ref) 'Ambrosetti Parameter
    k = 1# - (Tan(deg2rad * (Abs(theta) / 2#)) ^ (1# / r))
    
    Ambrosetti = k
End Function

Private Function Interpolation(numberOfRefs As Integer, theta_ref() As Double, K_ref() As Double, theta As Double)
    Dim k As Double
    Dim i As Integer
    
    i = 0
    
    Do While theta_ref(i) < Abs(theta)
        i = i + 1
    Loop
    
    k = K_ref(i - 1) + (K_ref(i) - K_ref(i - 1)) * ((Abs(theta) - theta_ref(i - 1)) / (theta_ref(i) - theta_ref(i - 1)))
    
    Interpolation = k
End Function

Private Function calcAvgEndLossEfficiency(theta As Double)
    Dim i As Integer
    Dim sumEffEndLoss As Double
    
    sumEffEndLoss = 0
    
    For i = 1 To n_series_coll
        sumEffEndLoss = sumEffEndLoss + calcEndLossEfficiencyIndividualCollector(i, theta)
    Next i
    
    calcAvgEndLossEfficiency = sumEffEndLoss / n_series_coll
End Function

Private Function calcEndLossEfficiencyIndividualCollector(collNumber As Integer, theta As Double)
    Dim effGains As Double
    Dim effLosses As Double
    Dim L_avg As Double
    
    L_avg = calcAvgSurfaceToFocusLength()
    
    If (collNumber Mod n_rows) = 0 Then
        effGains = 0
    Else
        effGains = (L_avg * Tan(deg2rad * theta) - distanceBetweenRows) / length_coll
    End If
    
    effLosses = 1 - L_avg * Tan(deg2rad * theta) / length_coll + effGains
    
    calcEndLossEfficiencyIndividualCollector = effLosses
End Function

Private Function calcAvgSurfaceToFocusLength()
    Dim L As Double
    Dim a As Double
    Dim halfw As Double
    
    a = focalLength
    halfw = width_coll / 2
    L = Sqr((4 * a ^ 2 + halfw ^ 2) ^ 2 / a ^ 2) * (12 * a ^ 2 + halfw ^ 2) / (12 * (4 * a ^ 2 + halfw ^ 2))
End Function

