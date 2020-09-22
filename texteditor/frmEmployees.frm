VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmEmployees 
   Caption         =   "Employees"
   ClientHeight    =   4245
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form3"
   ScaleHeight     =   4245
   ScaleWidth      =   5520
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5520
      TabIndex        =   1
      Top             =   3600
      Width           =   5520
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   4505
         TabIndex        =   6
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   300
         Left            =   3409
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   2313
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   300
         Left            =   1217
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   121
         TabIndex        =   2
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Data datPrimaryRS 
      Align           =   2  'Align Bottom
      Caption         =   " "
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\DevStudio\VB\Nwind.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmEmployees.frx":0000
      Top             =   3900
      Width           =   5520
   End
   Begin MSDBGrid.DBGrid grdDataGrid 
      Align           =   1  'Align Top
      Bindings        =   "frmEmployees.frx":00E3
      Height          =   3495
      Left            =   0
      OleObjectBlob   =   "frmEmployees.frx":023D
      TabIndex        =   0
      Top             =   0
      Width           =   5520
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
  datPrimaryRS.Recordset.MoveLast
  grdDataGrid.SetFocus
  SendKeys "{down}"
End Sub

Private Sub cmdDelete_Click()
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
End Sub

Private Sub cmdRefresh_Click()
  datPrimaryRS.Refresh
End Sub

Private Sub cmdUpdate_Click()
  datPrimaryRS.UpdateRecord
  datPrimaryRS.Recordset.Bookmark = datPrimaryRS.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub datPrimaryRS_Error(DataErr As Integer, Response As Integer)
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0
End Sub

Private Sub datPrimaryRS_Reposition()
  Screen.MousePointer = vbDefault
  datPrimaryRS.Caption = "Record: " & (datPrimaryRS.Recordset.AbsolutePosition + 1)
End Sub

Private Sub datPrimaryRS_Validate(Action As Integer, Save As Integer)
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
      Screen.MousePointer = vbDefault
  End Select
  Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  grdDataGrid.Height = Me.ScaleHeight - datPrimaryRS.Height - picButtons.Height - 30
End Sub

