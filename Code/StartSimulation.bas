Attribute VB_Name = "StartSimulation"
Option Explicit

Sub StartSolGainSimulation()
Attribute StartSolGainSimulation.VB_Description = "This Macro intialises a SolGain simulation by checking whether weather data has been provided and then opening the Geographical Inputs userform."
Attribute StartSolGainSimulation.VB_ProcData.VB_Invoke_Func = " \n14"
'
' StartSolGainSimulation Macro
' This Macro intialises a SolGain simulation by checking whether weather data has been provided and then opening the Geographical Inputs userform.
    
    Dim weatherWS As Worksheet
    
    Set weatherWS = ThisWorkbook.Sheets("Weather Data")
    
    If WorksheetFunction.CountBlank(weatherWS.Range("D4:D8763")) > 0 Then
        MsgBox "Please enter Diffuse Radiation on Horizontal Surface for every hour of the year (8760 hours in a year)"
        Exit Sub
    ElseIf WorksheetFunction.CountBlank(weatherWS.Range("E4:E8763")) > 0 Then
        MsgBox "Please enter Direct Normal Irradiation for every hour of the year (8760 hours in a year)"
        Exit Sub
    ElseIf WorksheetFunction.CountBlank(weatherWS.Range("F4:F8763")) > 0 Then
        MsgBox "Please enter Temperature Outside for every hour of the year (8760 hours in a year)"
        Exit Sub
    ElseIf WorksheetFunction.Count(weatherWS.Range("D4:F8763")) - 3 * 8760 <> 0 Then
        MsgBox "Please make sure all weather data entries are numbers"
        Exit Sub
    End If
    
    UserFormGeoInputs.Show
End Sub
