VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormDailyDemand 
   Caption         =   "Daily Demand"
   ClientHeight    =   6708
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15348
   OleObjectBlob   =   "UserFormDailyDemand.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormDailyDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HeatDemandWS As Worksheet

Private Sub CommandButtonNext_Click()
    UserFormDailyDemand.Hide
    UserFormWeeklyDemand.Show
End Sub

Private Sub CommandButtonPrev_Click()
    UserFormDailyDemand.Hide
    UserFormProcessInfo.Show
End Sub

Private Sub TextBox1_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox10_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox11_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub


Private Sub TextBox12_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox13_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox14_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox15_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox16_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox17_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox18_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox19_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox2_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox20_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox21_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox22_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox23_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox24_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox3_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub


Private Sub TextBox4_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox5_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox6_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox7_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox8_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBox9_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Set HeatDemandWS = ThisWorkbook.Sheets("Heat Demand Profile")
    
    If doesFolderExist(ThisWorkbook.Path & "\ProgramFiles") = False Then
        MkDir ThisWorkbook.Path & "\ProgramFiles"
    End If
    
    With UserFormDailyDemand
        For i = 1 To 24
            Controls("TextBoxHour" & i).Text = i
            Controls("TextBoxHour" & i).Locked = True
            Controls("TextBoxHour" & i).Enabled = False
            Controls("TextBox" & i).Text = 100
        Next i
    End With
    
    Call UpdateCells
End Sub

Private Sub UpdateCells()
    Dim i As Integer
    
    For i = 1 To 24

        If UserFormDailyDemand.Controls("TextBox" & i) <> "" Then
            HeatDemandWS.Cells(i + 2, 3) = UserFormDailyDemand.Controls("TextBox" & i) / 100#
        Else
            HeatDemandWS.Cells(i + 2, 3) = 0
        End If
    Next i
End Sub

Private Sub SaveChartUploadChart()
    Dim DemandChart As Chart
    Dim PathName As String
    
    'Save Chart to ProgramFiles folder
    Set DemandChart = HeatDemandWS.ChartObjects(1).Chart
    PathName = ThisWorkbook.Path & "\ProgramFiles\DailyProfile.jpg"
    DemandChart.Export FileName:=PathName ', FilterName:="JPEG"
    
    'Upload Chart to UserForm
    ImageDemandChart.Picture = LoadPicture(PathName)
End Sub

Private Function doesFolderExist(folderPath) As Boolean

    doesFolderExist = Dir(folderPath, vbDirectory) <> ""
    
End Function
