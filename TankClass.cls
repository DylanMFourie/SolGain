VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TankClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Tank characterization Declarations
Dim V As Double  'Volume of tank [m^3]
Dim U As Double  'Specific heat transfer coeff. of tank to environment [W/(m^2*K)]
Dim N_nodes As Integer   'Number of nodes used by tank model

'In object variables
Dim a As Double 'Surface area of outside of Tank [m^2]
Dim r As Double 'Tank radius [m]
Dim h As Double 'Tank height [m]
Dim T_tank() As Double
Dim T_tankNew() As Double
Dim T_tankNewPrevIt() As Double
Dim Fl() As Integer
Dim Fc() As Integer

'Tank outputs
Public T_toHE As Double
Public T_toProcess As Double
Public T_node3 As Double
Public T_node8 As Double

Sub constructor(V_tank_m3 As Double, U_tank As Double, N_tankNodes As Integer, T As TemperaturesClass)
    Dim i As Integer
    
    'If there is storage
    If V_tank_m3 > 0 Then
        V = V_tank_m3
        U = U_tank
        N_nodes = N_tankNodes
        
        'Calculate Tank radius, height and surface area
        r = (V / (2 * Pi)) ^ (1 / 3)
        h = 2 * r
        a = (2 * Pi * r * h) + (2 * Pi * r ^ 2)
        
        'Redimension arrays
        ReDim T_tank(1 To N_nodes)
        ReDim T_tankNew(1 To N_nodes)
        ReDim T_tankNewPrevIt(1 To N_nodes)
        ReDim SumT_tank(1 To N_nodes)
        ReDim Fl(1 To N_nodes)
        ReDim Fc(1 To N_nodes)
        
        'Set each node of Tank to ambient temp
        For i = 1 To N_nodes
            T_tank(i) = T.ambientTank_K
        Next i
        
        'Set each node of TankPrev to ambient temp
        T_tankNew = T_tank
        T_tankNewPrevIt = T_tank
        
        
        'Set
        T_toProcess = T_tank(1)
        T_toHE = T_tank(N_nodes)
        T_node3 = T_tank(3)
        T_node8 = T_tank(8)
    End If


    'ThisWorkbook.Sheets("Tank").Range("A1:L8900").Clear
End Sub

Sub calcOutputTempsWholeTimeStep(T As TemperaturesClass, tankHTF As HTFPropertiesClass, M_coll As Double, M_pro As Double, no As Integer)
    Dim n As Integer
    Dim numTimeSteps As Integer
    
    numTimeSteps = 6

    For n = 1 To numTimeSteps
        Call calcOutputTempsSmallTimeStep(T, tankHTF, M_coll, M_pro, numTimeSteps, no)
    Next n
    
    'Assign tank outputs
'    T_toHE = SumT_toHE / numTimeSteps
'    T_toProcess = SumT_toPro / numTimeSteps
    T_toHE = T_tank(N_nodes)
    T_toProcess = T_tank(1)
    T_node3 = T_tank(3)
    T_node8 = T_tank(8)
End Sub


Sub calcOutputTempsSmallTimeStep(T As TemperaturesClass, tankHTF As HTFPropertiesClass, M_coll As Double, M_pro As Double, numTimeSteps As Integer, no As Integer)
    Dim SumFl As Integer
    Dim SumFc As Integer
    Dim M_i As Double
    Dim M_iplus1 As Double
    Dim dTi As Double
    Dim dTi_1 As Double
    Dim i As Integer
    Dim iplus1 As Integer
    Dim j As Integer
    Dim MassOfNode As Double
    Dim A_node As Double
    Dim c_p As Double
    Dim x_i As Double
    Dim y_i As Double
'    Dim gamma_i As Double
    Dim a_i As Double
    Dim b_i As Double
    Dim c_i As Double
    Dim d_i As Double
    Dim itError As Double
    
    MassOfNode = tankHTF.rho * (V / N_nodes)
    
    Call calcControlFunctions(T)
    itError = 1000
    
    Do While itError > iterx
        T_tankNewPrevIt = T_tankNew
        
        'Enter loop to determine node Temperatures
        For i = 1 To N_nodes
            M_i = 0
            M_iplus1 = 0
            
            'Calculate mass flow between nodes i and i-1
            
            If i <> 1 Then
                SumFc = 0
                For j = 1 To (i - 1)
                    SumFc = SumFc + Fc(j)
                Next j
                
                SumFl = 0
                For j = (i) To N_nodes
                    SumFl = SumFl + Fl(j)
                Next j
                
                M_i = M_coll * SumFc - M_pro * SumFl
            End If
            
            'Calculate mass flow between nodes i+1 and i
            iplus1 = i + 1   '**Kreef i
            If iplus1 <> (N_nodes + 1) Then
                SumFc = 0
                For j = 1 To (iplus1 - 1)
                    SumFc = SumFc + Fc(j)
                Next j
                
                SumFl = 0
                For j = (iplus1) To N_nodes
                    SumFl = SumFl + Fl(j)
                Next j
                
                M_iplus1 = M_coll * SumFc - M_pro * SumFl
            End If
            
            'Calculate Surface Area of Node Exposed to Ambient Temp
            If i = 1 Or i = N_nodes Then
                A_node = (2 * Pi * r * (h / N_nodes)) + (Pi * r ^ 2)
            Else
                A_node = (2 * Pi * r * (h / N_nodes))
            End If
            
            'Assign value of c_p
            c_p = tankHTF.c_p
            
            If M_i > 0 Then
                x_i = M_i
            Else
                x_i = 0
            End If

            If M_iplus1 < 0 Then
                y_i = M_iplus1
            Else
                y_i = 0
            End If

'            If M_i > 0 Then
'                gamma_i = 1
'            Else
'                gamma_i = 0
'            End If
            
            a_i = x_i / MassOfNode
            b_i = (-(U * A_node / c_p) - Fc(i) * M_coll - Fl(i) * M_pro - x_i + y_i) / MassOfNode
            c_i = -y_i / MassOfNode
            d_i = ((U * A_node / c_p) * T.ambientTank_K + Fc(i) * M_coll * T.HEToTank_K + Fl(i) * M_pro * T.processOut_K) / MassOfNode
            
            If i = 1 Then
                dTi_1 = (b_i / 2) * (T_tank(i) + T_tankNew(i)) + (c_i / 2) * (T_tank(i + 1) + T_tankNew(i + 1)) + d_i
            ElseIf i = N_nodes Then
                dTi_1 = (a_i / 2) * (T_tank(i - 1) + T_tankNew(i - 1)) + (b_i / 2) * (T_tank(i) + T_tankNew(i)) + d_i
            Else
                dTi_1 = (a_i / 2) * (T_tank(i - 1) + T_tankNew(i - 1)) + (b_i / 2) * (T_tank(i) + T_tankNew(i)) + (c_i / 2) * (T_tank(i + 1) + T_tankNew(i + 1)) + d_i
            End If
            
            T_tankNew(i) = T_tank(i) + dTi_1 * (timeStep / numTimeSteps)
            'ThisWorkbook.Sheets("Tank").Cells(no + 1, i) = Abs(dTi - dTi_1)
        Next i
        itError = calcError()
    Loop
    
    'Assign previous Tank Profile
    T_tank = T_tankNew
End Sub

Private Sub calcControlFunctions(T As TemperaturesClass)
    Dim i As Integer
    
    'Calculate control functions to determine where water returning from coll and load goes to
    For i = 1 To N_nodes
    
        'Initialize load and collector control functions to 0
        Fl(i) = 0
        Fc(i) = 0
        
        'Determine value of collector control function
        If i = 1 Then
            If T.HEToTank_K > T_tank(i) Then
                Fc(i) = 1
            End If
        ElseIf i = N_nodes And T.HEToTank_K < T_tank(i) Then
            Fc(i) = 1
        ElseIf T_tank(i - 1) >= T.HEToTank_K And T.HEToTank_K > T_tank(i) Then
            Fc(i) = 1
        End If
        
        'Determine value of load control function
        If i = 1 Then
            If T.processOut_K > T_tank(i) Then
                Fl(i) = 1
            End If
        ElseIf i = N_nodes And T.processOut_K < T_tank(i) Then
            Fl(i) = 1
        ElseIf T_tank(i - 1) >= T.processOut_K And T.processOut_K > T_tank(i) Then
            Fl(i) = 1
        End If
        
    Next i
End Sub

Sub displayTempLevels(i As Integer, firstCol As Integer)
    Dim j As Integer
    
    For j = 1 To N_nodes
        ThisWorkbook.Sheets("Tester").Cells(i + 1, j + (firstCol - 1)) = T_tank(j)
    Next j
End Sub

Private Function MpoptCalc(Q_need, c_p, T_toProcess, T_fromProcess)
    MpoptCalc = Q_need / (c_p * (T_toProcess - T_fromProcess))
End Function

Private Function calcError()
    Dim i As Integer
    Dim error_i As Double
    Dim error_max As Double
    
    For i = 1 To N_nodes
        error_i = Abs(T_tankNew(i) - T_tankNewPrevIt(i))
        
        If error_max < error_i Then
            error_max = error_i
        End If
    Next i
End Function
