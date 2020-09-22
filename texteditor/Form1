VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSChart"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin MSChartLib.MSChart MSChart1 
      Height          =   6015
      Left            =   0
      OleObjectBlob   =   "Form1.frx":0442
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim a, m, b, r, c, d, e, q, l As Double
Dim x, Y, w, u, h, j As String
r = InputBox("Enter number of rows(INTEGER)", a)
On Error Resume Next
c = InputBox("Enter number of coloumns(INTEGER)", b)
On Error Resume Next
l = InputBox("Enter no. of Coloumn Labels(INTEGER)", q)
On Error Resume Next
x = InputBox("Enter Chart Type(3D,2D,Line,Pie,Area)", Y)
On Error Resume Next
With Form1.MSChart1
On Error Resume Next
If x = "3D" Then
        .chartType = VtChChartType3dBar
Else
If x = "2D" Then
        .chartType = VtChChartType2dBar
Else
If x = "Line" Then
        .chartType = VtChChartType2dLine
Else
If x = "Pie" Then
        .chartType = VtChChartType2dPie
Else
If x = "Area" Then
        .chartType = VtChChartType2dArea
Else
End If
End If
End If
End If
End If
        .ColumnCount = c
        .RowCount = r
            For Column = 1 To c
            w = InputBox("Enter Coloumn Label", u)
            j = InputBox("Enter Row Label(Along X-Axis)", h)
              For Row = 1 To r
                .Column = Column
                .Row = Row
                .ColumnLabel = w
                .RowLabel = j
                .Data = InputBox("Enter value (Along Y-Axis)", e)
                Next Row
              Next Column
 On Error Resume Next
        .ShowLegend = True
        .SelectPart VtChPartTypePlot, index1, index2, index3, index4
        .EditCopy
        .SelectPart VtChPartTypeLegend, index1, index2, index3, index4
        .EditPaste
    End With
On Error Resume Next
End Sub


