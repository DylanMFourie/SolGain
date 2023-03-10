VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormWeeklyDemand 
   Caption         =   "Weekly Demand"
   ClientHeight    =   6924
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12804
   OleObjectBlob   =   "UserFormWeeklyDemand.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormWeeklyDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HeatDemandWS As Worksheet

Private Sub CommandButtonNext_Click()
    UserFormWeeklyDemand.Hide
    UserFormYearlyDemand.Show
End Sub

Private Sub CommandButtonPrev_Click()
    UserFormWeeklyDemand.Hide
    UserFormDailyDemand.Show
End Sub

Private Sub TextBoxPercent1_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent2_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent3_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent4_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent5_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent6_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub TextBoxPercent7_AfterUpdate()
    Call UpdateCells
    Call SaveChartUploadChart
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Set HeatDemandWS = ThisWorkbook.Sheets("Heat Demand Profile")
    
    'Setup Directory folder
    If doesFolderExist(ThisWorkbook.Path & "\ProgramFiles") = False Then
        MkDir ThisWorkbook.Path & "\ProgramFiles"
    End If
    
    'Setup weekday textboxes
    TextBoxHour1.Text = "Mon"
    TextBoxHour2.Text = "Tue"
    TextBoxHour3.Text = "Wed"
    TextBoxHour4.Text = "Thu"
    TextBoxHour5.Text = "Fri"
    TextBoxHour6.Text = "Sat"
    TextBoxHour7.Text = "Sun"
    
    With UserFormWeeklyDemand
        For i = 1 To 7
            Controls("TextBoxHour" & i).Locked = True
            Controls("TextBoxHour" & i).Enabled = False
            Controls("TextBoxPercent" & i).Text = 100
        Next i
    End With
    
End Sub

Private Sub UpdateCells()
    Dim i As Integer
    
    For i = 1 To 7
        If UserFormWeeklyDemand.Controls("TextBoxPercent" & i) <> "" Then
            HeatDemandWS.Cells(i + 2, 6) = UserFormWeeklyDemand.Controls("TextBoxPercent" & i) / 100#
        Else
            HeatDemandWS.Cells(i + 2, 6) = 0
        End If
    Next i
End Sub

Private Sub SaveChartUploadChart()
    Dim DemandChart As Chart
    Dim PathName As String
    
    'Save Chart to ProgramFiles folder
    Set DemandChart = HeatDemandWS.ChartObjects(2).Chart
    PathName = ThisWorkbook.Path & "\ProgramFiles\WeeklyProfile.jpg"
    DemandChart.Export FileName:=PathName ', FilterName:="JPEG"
    
    'Upload Chart to UserForm
    ImageDemandChart.Picture = LoadPicture(PathName)
End Sub

Private Function doesFolderExist(folderPath) As Boolean

    doesFolderExist = Dir(folderPath, vbDirectory) <> ""
    
End Function

