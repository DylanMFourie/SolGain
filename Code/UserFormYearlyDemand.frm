VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormYearlyDemand 
   Caption         =   "Yearly Demand"
   ClientHeight    =   10524
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17148
   OleObjectBlob   =   "UserFormYearlyDemand.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormYearlyDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HeatDemandWS As Worksheet



Private Sub CommandButtonNext_Click()
    UserFormYearlyDemand.Hide
    UserFormFirstCollector.Show
End Sub

Private Sub CommandButtonPrev_Click()
    UserFormYearlyDemand.Hide
    UserFormWeeklyDemand.Show
End Sub

Private Sub TextBoxPercent1_AfterUpdate()
    Call UpdateCells '(1)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent10_AfterUpdate()
    Call UpdateCells '(10)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent11_AfterUpdate()
    Call UpdateCells '(11)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent12_AfterUpdate()
    Call UpdateCells '(12)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent13_AfterUpdate()
    Call UpdateCells '(13)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent14_AfterUpdate()
    Call UpdateCells '(14)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent15_AfterUpdate()
    Call UpdateCells '(15)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent16_AfterUpdate()
    Call UpdateCells '(16)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent17_AfterUpdate()
    Call UpdateCells '(17)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent18_AfterUpdate()
    Call UpdateCells '(18)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent19_AfterUpdate()
    Call UpdateCells '(19)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent2_AfterUpdate()
    Call UpdateCells '(2)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent20_AfterUpdate()
    Call UpdateCells '(20)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent21_AfterUpdate()
    Call UpdateCells '(21)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent22_AfterUpdate()
    Call UpdateCells '(22)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent23_AfterUpdate()
    Call UpdateCells '(23)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent24_AfterUpdate()
    Call UpdateCells '(24)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent25_AfterUpdate()
    Call UpdateCells '(25)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent26_AfterUpdate()
    Call UpdateCells '(26)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent27_AfterUpdate()
    Call UpdateCells '(27)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent28_AfterUpdate()
    Call UpdateCells '(28)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent29_AfterUpdate()
    Call UpdateCells '(29)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent3_AfterUpdate()
    Call UpdateCells '(3)
    Call SaveChartUploadChart
End Sub


Private Sub TextBoxPercent30_AfterUpdate()
    Call UpdateCells '(30)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent31_AfterUpdate()
    Call UpdateCells '(31)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent32_AfterUpdate()
    Call UpdateCells '(32)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent33_AfterUpdate()
    Call UpdateCells '(33)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent34_AfterUpdate()
    Call UpdateCells '(34)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent35_AfterUpdate()
    Call UpdateCells '(35)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent36_AfterUpdate()
    Call UpdateCells '(36)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent37_AfterUpdate()
    Call UpdateCells '(37)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent38_AfterUpdate()
    Call UpdateCells '(38)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent39_AfterUpdate()
    Call UpdateCells '(39)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent4_AfterUpdate()
    Call UpdateCells '(4)
    Call SaveChartUploadChart
End Sub


Private Sub TextBoxPercent40_AfterUpdate()
    Call UpdateCells '(40)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent41_AfterUpdate()
    Call UpdateCells '(41)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent42_AfterUpdate()
    Call UpdateCells '(42)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent43_AfterUpdate()
    Call UpdateCells '(43)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent44_AfterUpdate()
    Call UpdateCells '(44)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent45_AfterUpdate()
    Call UpdateCells '(45)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent46_AfterUpdate()
    Call UpdateCells '(46)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent47_AfterUpdate()
    Call UpdateCells '(47)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent48_AfterUpdate()
    Call UpdateCells '(48)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent49_AfterUpdate()
    Call UpdateCells '(49)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent5_AfterUpdate()
    Call UpdateCells '(5)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent50_AfterUpdate()
    Call UpdateCells '(50)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent51_AfterUpdate()
    Call UpdateCells '(51)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent52_AfterUpdate()
    Call UpdateCells '(52)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent53_AfterUpdate()
    Call UpdateCells '(53)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent6_AfterUpdate()
    Call UpdateCells '(6)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent7_AfterUpdate()
    Call UpdateCells '(7)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent8_AfterUpdate()
    Call UpdateCells '(8)
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent9_AfterUpdate()
    Call UpdateCells '(9)
    Call SaveChartUploadChart
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Set HeatDemandWS = ThisWorkbook.Sheets("Heat Demand Profile")
    
    If doesFolderExist(ThisWorkbook.Path & "\ProgramFiles") = False Then
        MkDir ThisWorkbook.Path & "\ProgramFiles"
    End If
    
    With UserFormYearlyDemand
        For i = 1 To 53
            Controls("TextBoxHour" & i).Text = i
            Controls("TextBoxHour" & i).Locked = True
            Controls("TextBoxHour" & i).Enabled = False
            Controls("TextBoxPercent" & i).Text = 100
        Next i
    End With
    
    Call UpdateCells
End Sub

Private Sub UpdateCells()
    Dim i As Integer
    
    For i = 1 To 53
        If UserFormYearlyDemand.Controls("TextBoxPercent" & i) <> "" Then
            HeatDemandWS.Cells(i + 2, 9) = UserFormYearlyDemand.Controls("TextBoxPercent" & i) / 100#
        Else
            HeatDemandWS.Cells(i + 2, 9) = 0
        End If
    Next i
End Sub

Private Sub SaveChartUploadChart()
    Dim DemandChart As Chart
    Dim PathName As String
    
    'Save Chart to ProgramFiles folder
    Set DemandChart = HeatDemandWS.ChartObjects(3).Chart
    PathName = ThisWorkbook.Path & "\ProgramFiles\YearlyProfile.jpg"
    DemandChart.Export FileName:=PathName ', FilterName:="JPEG"
    
    'Upload Chart to UserForm
    ImageDemandChart.Picture = LoadPicture(PathName)
End Sub

Private Function doesFolderExist(folderPath) As Boolean
    
    doesFolderExist = Dir(folderPath, vbDirectory) <> ""
    
End Function

