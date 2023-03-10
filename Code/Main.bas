Attribute VB_Name = "Main"
Option Explicit

Public Const Pi As Double = 3.14159265358979
Public Const deg2rad As Double = 1.74532925199433E-02  'radian
Public Const timeStep As Double = (60 * 60) 'seconds
Public Const iterx As Double = 0.00001
Public Const kelvin0 As Double = 273.15

Public Const White As Long = &H80000005
Public Const Grey As Long = &HE0E0E0

Public Const ProHeat As String = "Non-Concentrating (Temps less than ~110?C)"
Public Const DirectSteam As String = "Concentrating (Direct steam generation T > 110?C)"
Public Const FPC As String = "Flate Plate Collector (T < 60degC)"
Public Const ETC As String = "Evacuated Tube Collector (T < 100degC)"
Public Const CPC As String = "Compound Parabolic Collector (T < 100degC)"
Public Const ParabolicTrough As String = "Parabolic Trough"
Public Const Water As String = "Water"
Public Const Glycol20 As String = "Glycol 20%"
Public Const Glycol30 As String = "Glycol 30%"
Public Const Glycol40 As String = "Glycol 40%"
Public Const Glycol42 As String = "Glycol 42%"
Public Const Glycol50 As String = "Glycol 50%"

Public Const Therminol66 As String = "Therminol 66"

Sub Simulation()
    'Declare non-physical objects
    Dim T As New TemperaturesClass  'All Temps in degree K
    Dim T_dummy As New TemperaturesClass    'All Temps in degree K
    Dim Tprev As New TemperaturesClass  'All Temps in degree K
    Dim collectorHTF As New HTFPropertiesClass
    Dim processHTF As New HTFPropertiesClass
    Dim massFlow As New MassFlowClass
    Dim finances As New FinancialModelClass
    
    'Declare physical objects
    Dim sun As New SunClass
    Dim collector As New CollectorClass
    Dim pipe As New PipeClass
    Dim heatExchanger As New HeatExchangerClass
    Dim Tank As New TankClass
    
    'Declare Main subroutine variables
    Dim i As Integer    'iteration counter
    Dim j As Integer    'iteration counter
    Dim k As Integer    'iteration counter
    Dim colError As Double   'iteration error
    Dim week As Integer
    Dim day As Integer
    Dim hour As Integer
    Dim Q_aux As Double
    Dim Q_auxNoSolar As Double
    Dim E_solarToProcess_kWh As Double
    Dim E_auxToProcess_kWh As Double
    Dim E_auxNoSolar_kWh As Double
    Dim I_sunWholeYear_kWh As Double
    
    'Declare Worksheet Objects
    Dim collInputWS As Worksheet
    Dim proInputWS As Worksheet
    Dim GeoInputWS As Worksheet
    Dim miscInputWS As Worksheet
    Dim weatherWS As Worksheet
    Dim demandWS As Worksheet
    Dim resultsWS As Worksheet
    Dim testerWS As Worksheet
    Dim financeWS As Worksheet
    
    'Initialize Worksheet Objects
    Set collInputWS = ThisWorkbook.Sheets("Collector Inputs")
    Set proInputWS = ThisWorkbook.Sheets("Process Inputs")
    Set GeoInputWS = ThisWorkbook.Sheets("Geographical Inputs")
    Set miscInputWS = ThisWorkbook.Sheets("Misc Inputs")
    Set weatherWS = ThisWorkbook.Sheets("Weather Data")
    Set demandWS = ThisWorkbook.Sheets("Heat Demand Profile")
    Set resultsWS = ThisWorkbook.Sheets("Results")
    Set testerWS = ThisWorkbook.Sheets("Tester")
    Set financeWS = ThisWorkbook.Sheets("Financials")
    
    'Variable to control loops based on auto or manual collector inputs
    Dim n_systems As Integer
    
    If collector.manualInputs = True Then
        n_systems = 1
    ElseIf collector.manualInputs = False Then
        n_systems = 5
    End If
    
    For k = 1 To 2
    
        For j = 1 To n_systems
        
            'Construct Non-physical Objects
            Call T.constructor(proInputWS.Range("A2"), proInputWS.Range("B2"), weatherWS.Range("F4"), 25)
            Call collectorHTF.constructor(1000, 4180)
            Call processHTF.constructor(1000, 4180)
            If k = 1 Then
                Call massFlow.constructor(proInputWS.Range("D2"), processHTF.c_p, proInputWS.Range("A2"), proInputWS.Range("B2"), "Pre Heating")
            Else
                Call massFlow.constructor(proInputWS.Range("D2"), processHTF.c_p, proInputWS.Range("A2"), proInputWS.Range("B2"), "Target Temperature")
            End If
            Call finances.constructor(0.03, 15, 0.08 * 15, 0.1)
            
            'Construct Physical Objects
            Call sun.constructor(GeoInputWS.Range("F2"), GeoInputWS.Range("B2"), GeoInputWS.Range("D2"), GeoInputWS.Range("E2"))
            Call collector.constructor(collInputWS.Range("K2"), collInputWS.Range("L2"), 1, j * 2, T.ambient_K)
            Call pipe.constructor(miscInputWS.Range("F2"), miscInputWS.Range("G2"), miscInputWS.Range("H2"))
            Call heatExchanger.constructor(miscInputWS.Range("E2"))
            Call Tank.constructor(0.2 * collector.grossApertureArea_field, miscInputWS.Range("D2"), 12, T)
            
            'Set incremental amount of energy provided by solar to zero
            E_solarToProcess_kWh = 0
            E_auxToProcess_kWh = 0
            E_auxNoSolar_kWh = 0
            I_sunWholeYear_kWh = 0
            finances.costDifference_R = 0
            
            'Clear contents
            testerWS.Range("A2:BZ8761").ClearContents
            
            'collector.n_parallel_modules = n * 2
            
            'Simulate every hour of the year
            For i = 1 To (24 * 365)
                'Calculate heat demand for this hour
                'Call heatDemand.calcHeatDemand(i)
                
                'Set HE inlet from tank and auxilary heater inlet (remains const. during time step)
                T.HEFromTank_K = Tank.T_toHE
                T.auxIn_K = Tank.T_toProcess
                If Tank.T_toProcess > T.processOut_K Then
                    T.auxIn_K = Tank.T_toProcess
                Else
                    T.auxIn_K = T.processOut_K
                End If
                
                'Set aux heater if no solar gains are made
                T.auxNoSolar_K = T.processOut_K
                
                'Sun position calculation
                Call sun.calcSunPos(i)
                
                'Collector tilt calculation
                Call collector.calcTilt(sun)
                
                'Solar radiation on tilted surface calculation
                Call sun.calcIrradianceComps(weatherWS.Cells(i + 3, 4), weatherWS.Cells(i + 3, 5), collector)
                
                Call collector.calcIAMs(sun)
                
                'Calculate Collector Mass Flow Rate
                If massFlow.heatingStrategy = "Target Temperature" Then
                    'Assign Known Temps
                    T.HEToTank_K = massFlow.T_processTarget_K
                    
                    colError = 100
                    Do While colError > iterx
                        massFlow.M_coll_prev = massFlow.M_collector
                        Set Tprev = T
                        
                        Call heatExchanger.calcHotSideTemps(T.HEFromTank_K, T.HEToTank_K, processHTF, collectorHTF, massFlow.M_collector, massFlow.M_collector)
                        T.HEFromCollector_K = heatExchanger.T_hotIn
                        T.HEToCollector_K = heatExchanger.T_hotOut
                        
                        T.collectorOut_K = pipe.calcInletTemp_K(T.HEFromCollector_K, T.ambient_K, collectorHTF, massFlow.M_collector)
                        T.collectorIn_K = pipe.calcOutletTemp_K(T.HEToCollector_K, T.ambient_K, collectorHTF, massFlow.M_collector)
                        
                        massFlow.M_collector = collector.approxFlowRate(sun, T.collectorIn_K, T.collectorOut_K, T.ambient_K, collectorHTF)
                        testerWS.Cells(i + 1, 18) = massFlow.M_collector
                        massFlow.M_collector = collector.calcFlowRate(sun, T.collectorIn_K, T.collectorOut_K, T.ambient_K, collectorHTF, massFlow.M_collector)
                        testerWS.Cells(i + 1, 19) = massFlow.M_collector
                        
                        colError = calcError(T, Tprev)
                        colError = WorksheetFunction.Max(colError, Abs(massFlow.M_collector - massFlow.M_coll_prev))
                    Loop
                ElseIf massFlow.heatingStrategy = "Pre Heating" Then
                    massFlow.M_collector = (collector.apertureArea_singleColl * collector.n_series_coll * collector.n_parallel_modules) * 18# / (60# * 60#)
                End If
                
                'iteratively find collector loop Temperatures
                colError = 100  'Arbitrary large number
                Do While colError > iterx
                    'Set Tprev to T
                    Set Tprev = T
                    
                    'Calculate collector output temperature
                    T.collectorOut_K = collector.calcOutputTemp(sun, T.collectorIn_K, T.ambient_K, collectorHTF, massFlow.M_collector)
                    
                    'Calculate collector side HE inlet temp
                    T.HEFromCollector_K = pipe.calcOutletTemp_K(T.collectorOut_K, T.ambient_K, collectorHTF, massFlow.M_collector)
                    
                    'Calculate collector side Heat Exchanger outlet temperature
                    Call heatExchanger.calcOutTemps(T.HEFromCollector_K, T.HEFromTank_K, collectorHTF, processHTF, massFlow.M_collector, massFlow.M_collector)
                    T.HEToCollector_K = heatExchanger.T_hotOut
                    
                    'Calculate collector inlet temperature
                    T.collectorIn_K = pipe.calcOutletTemp_K(T.HEToCollector_K, T.ambient_K, collectorHTF, massFlow.M_collector)
                    
                    'Calc Error
                    colError = calcError(T, Tprev)
                Loop
                
                'Calculate Temp into Tank
                T.HEToTank_K = heatExchanger.T_coldOut
                
                'Determine which week, day and hour it is
                week = Int((i - 1) / 168) + 1
                day = Int((i - 1) / 24) - (week - 1) * 7 + 1
                hour = (i) - (week - 1) * 168 - (day - 1) * 24
                
                'Calculate Process Mass Flow Rate
                massFlow.M_process = massFlow.calcProcessFlowRate(demandWS.Range("C" & hour + 2), demandWS.Range("F" & day + 2), demandWS.Range("I" & week + 2))
                
                'Calculate heat input from Auxiliary Burner
                Q_aux = massFlow.M_process * processHTF.c_p * (T.processIn_K - T.auxIn_K)
                
                'calculate energy supplied by auxilary heat for the year
                E_auxToProcess_kWh = E_auxToProcess_kWh + 1 * Q_aux / 1000
                
                'Calculate energy supplied by solar for the year
                E_solarToProcess_kWh = E_solarToProcess_kWh + 1 * massFlow.M_process * processHTF.c_p * (T.auxIn_K - T.processOut_K) / 1000
                
                'Calculate all energy supplied by the sun
                I_sunWholeYear_kWh = I_sunWholeYear_kWh + 1 * (collector.n_series_coll * collector.n_parallel_modules * collector.apertureArea_singleColl) * (sun.G_bt + sun.G_rt + sun.G_st) / 1000
                
                'Calculate heat input from Heater if there are no solar gains
                Q_auxNoSolar = massFlow.M_process * processHTF.c_p * (T.processIn_K - T.auxNoSolar_K)
                
                'Calculate energy supplied by aux if no solar heat provided
                E_auxNoSolar_kWh = E_auxNoSolar_kWh + 1 * Q_auxNoSolar / 1000
                
                'Display Results in tester
                testerWS.Cells(i + 1, 1) = T.collectorIn_K - kelvin0
                testerWS.Cells(i + 1, 4) = T.collectorOut_K - kelvin0
                testerWS.Cells(i + 1, 7) = T.HEFromCollector_K - kelvin0
                testerWS.Cells(i + 1, 10) = T.HEToCollector_K - kelvin0
                testerWS.Cells(i + 1, 13) = T.HEToTank_K - kelvin0
                testerWS.Cells(i + 1, 16) = T.HEFromTank_K - kelvin0
                testerWS.Cells(i + 1, 17) = T.auxIn_K - kelvin0
                testerWS.Cells(i + 1, 19) = massFlow.M_collector
                testerWS.Cells(i + 1, 20) = massFlow.M_process
                testerWS.Cells(i + 1, 22) = Q_aux
                testerWS.Cells(i + 1, 24) = finances.costDifference_R
                'Call Tank.displayTempLevels(i, 15)
                
                'Update Tank Profile
                If T.auxIn_K = T.processOut_K Then
                    Call Tank.calcOutputTempsWholeTimeStep(T, processHTF, massFlow.M_collector, 0, i)
                Else
                    Call Tank.calcOutputTempsWholeTimeStep(T, processHTF, massFlow.M_collector, massFlow.M_process, i)
                End If
                
                'Update ambient Temperatures
                Call T.updateAmbientTemp(weatherWS.Cells(i + 3, 6))
            Next i
            
            'Determine cost difference
            Call finances.CostDifference_SPHvsHeater(E_auxNoSolar_kWh, E_auxToProcess_kWh)
            
            'Determine solar fraction
            Call finances.calcSolarFraction(E_solarToProcess_kWh, E_auxToProcess_kWh)
            
            'Determine system efficiency
            Call finances.calcSystemEfficiency(E_solarToProcess_kWh, I_sunWholeYear_kWh)
            
            'Calculate system investment cost
            Call finances.calcSystemCost_R(collector.grossApertureArea_field)
            'finances.systemCost_R = 2000
                 
            'Calculate Net Present Value
            'Call finances.calcNPV
            
            'Calculate IRR
            Call finances.calcIRR_and_NPV
            
            'Calculate LCOH
            Call finances.calcLCOH_R(E_solarToProcess_kWh)
            
            Dim spacing As Integer
            
            If k = 1 Then
                spacing = 1 + j
            Else
                spacing = 9 + j
            End If
            
            financeWS.Cells(3, spacing) = collector.n_parallel_modules * collector.n_series_coll * collector.apertureArea_singleColl
            financeWS.Cells(4, spacing) = finances.solarFraction
            financeWS.Cells(5, spacing) = finances.systemEfficiency
            financeWS.Cells(6, spacing) = finances.NPV_R
            financeWS.Cells(7, spacing) = finances.IRR
            financeWS.Cells(8, spacing) = finances.LCOH_RperKWh
        Next j
    Next k
    
End Sub

Private Function calcError(Tcurrent As TemperaturesClass, Tprevious As TemperaturesClass)
    Dim itError(3) As Double
    
    itError(0) = Tcurrent.collectorIn_K - Tprevious.collectorIn_K
    itError(1) = Tcurrent.collectorOut_K - Tprevious.collectorOut_K
    itError(2) = Tcurrent.HEFromCollector_K - Tprevious.HEFromCollector_K
    itError(3) = Tcurrent.HEToCollector_K - Tprevious.HEToCollector_K
    
    calcError = WorksheetFunction.Max(itError)
End Function

Private Function calculateNoCollectorsInArea(availableArea As Double, collectorArea As Double)
    Dim noCols As Integer
    
    noCols = WorksheetFunction.RoundDown(availableArea / collectorArea, 0)
    
    calculateNoCollectorsInArea = noCols
End Function
