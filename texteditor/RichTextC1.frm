VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "midfile"; "mid"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RichTextC 
   Caption         =   "Rich Text Document"
   ClientHeight    =   10110
   ClientLeft      =   60
   ClientTop       =   870
   ClientWidth     =   15240
   Icon            =   "RichTextC1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10110
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command24 
      Caption         =   "Table"
      Height          =   255
      Left            =   11160
      TabIndex        =   35
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Cal"
      Height          =   255
      Left            =   10560
      TabIndex        =   34
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command21 
      Height          =   375
      Left            =   12960
      Picture         =   "RichTextC1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Maintain employee data"
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Toolbar1 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   10485
      TabIndex        =   15
      Top             =   0
      Width           =   10545
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   120
         Picture         =   "RichTextC1.frx":0654
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   720
         Picture         =   "RichTextC1.frx":0756
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   1320
         Picture         =   "RichTextC1.frx":0858
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Normal"
         Height          =   375
         Left            =   1920
         Picture         =   "RichTextC1.frx":095A
         TabIndex        =   28
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   2640
         Picture         =   "RichTextC1.frx":0E8C
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   3360
         Picture         =   "RichTextC1.frx":0F8E
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Bullets"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Font"
         Height          =   375
         Left            =   4080
         TabIndex        =   25
         ToolTipText     =   "Select a Font"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Height          =   375
         Left            =   4800
         Picture         =   "RichTextC1.frx":1090
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Open File"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Height          =   375
         Left            =   5520
         Picture         =   "RichTextC1.frx":1192
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Save As..."
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command11 
         Height          =   375
         Left            =   6240
         Picture         =   "RichTextC1.frx":1294
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Delete File"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command12 
         Height          =   375
         Left            =   6960
         Picture         =   "RichTextC1.frx":141E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Print"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command13 
         Height          =   375
         Left            =   8280
         Picture         =   "RichTextC1.frx":1520
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Center Text"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command14 
         Height          =   375
         Left            =   7680
         Picture         =   "RichTextC1.frx":1622
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Align Left"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command15 
         Height          =   375
         Left            =   8880
         Picture         =   "RichTextC1.frx":1724
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Align Right"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command16 
         Height          =   375
         Left            =   9480
         Picture         =   "RichTextC1.frx":1826
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Apply Colors"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command17 
         Height          =   375
         Left            =   9960
         Picture         =   "RichTextC1.frx":1D58
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Search"
         Top             =   0
         Width           =   495
      End
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   840
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   873
      _Version        =   327682
      BorderStyle     =   1
   End
   Begin VB.CommandButton Command19 
      Height          =   375
      Left            =   12360
      Picture         =   "RichTextC1.frx":1E5A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Insert Image"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command20 
      Height          =   375
      Left            =   11760
      Picture         =   "RichTextC1.frx":1F5C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Charts"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command18 
      Height          =   375
      Left            =   11160
      Picture         =   "RichTextC1.frx":239E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Internet"
      Top             =   0
      Width           =   495
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   13680
      LinkItem        =   "StatusBar1"
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   0
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   6480
      Top             =   6120
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   315
      ScaleWidth      =   15180
      TabIndex        =   8
      Top             =   9735
      Width           =   15240
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   4440
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6120
      TabIndex        =   2
      Top             =   4080
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   6120
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   4320
      MouseIcon       =   "RichTextC1.frx":27E0
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1320
      Width           =   4815
      Begin VB.TextBox Text1 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "RichTextC1.frx":2C22
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "A&pply"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   3360
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Font Size:"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15266
      _Version        =   327680
      Enabled         =   -1  'True
      ScrollBars      =   3
      MousePointer    =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"RichTextC1.frx":2C2F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   375
      Left            =   10560
      OleObjectBlob   =   "RichTextC1.frx":2CF8
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\My Documents\aman docs\GEEculator1.23.exe"
      TabIndex        =   9
      Top             =   0
      Width           =   495
   End
   Begin midfileCtl.midfile midfile1 
      Height          =   0
      Left            =   0
      OleObjectBlob   =   "RichTextC1.frx":37510
      TabIndex        =   33
      Top             =   0
      Width           =   0
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu CUT 
         Caption         =   "&Cut                                  Ctrl+X"
      End
      Begin VB.Menu COPY 
         Caption         =   "&Copy                               Ctrl+C"
      End
      Begin VB.Menu PASTE 
         Caption         =   "&Paste                              Ctrl+V"
      End
      Begin VB.Menu CLOSE 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu About 
      Caption         =   "&About"
      Index           =   1
      Begin VB.Menu About1 
         Caption         =   "About"
         Index           =   2
      End
   End
End
Attribute VB_Name = "RichTextC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About1_Click(Index As Integer)
Load frmAbout
frmAbout.Visible = True
End Sub

Private Sub CLOSE_Click()
Unload Me
Set richtextc1 = Nothing
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo2_Change()
If Combo2.Text = "Times New Roman" Then
RichTextBox1.Font = "Times New Roman"
Text1.Font = "Times New Roman"
Else
End If
If Combo2.Text = "MS" Then
RichTextBox1.Font = "MS Sans Serif"
Text1.Font = "MS Sans Serif"
Else
End If
If Combo2.Text = "Arial" Then
RichTextBox1.Font = "Arial"
Text1.Font = "Arial"
Else
End If
If Combo2.Text = "Bookman Old Style" Then
RichTextBox1.Font = "Bookman Old Style"
Text1.Font = "Bookman Old Style"
Else
End If
If Combo2.Text = "Century" Then
RichTextBox1.Font = "Century"
Text1.Font = "Century"
Else
End If
If Combo2.Text = "Arial Black" Then
RichTextBox1.Font = "Arial Black"
Text1.Font = "Arial Black"
Else
End If
If Combo2.Text = "Century Gothic" Then
RichTextBox1.Font = "Century Gothic"
Text1.Font = "Century Gothic"
Else
End If
If Combo2.Text = "Impact" Then
RichTextBox1.Font = "Impact"
Text1.Font = "Impact"
Else
End If
If Combo2.Text = "Bookman Antiqua" Then
RichTextBox1.Font = "Bookman Antiqua"
Text1.Font = "Bookman Antiqua"
Else
End If
If Combo2.Text = "Comic Sans MS" Then
RichTextBox1.Font = "Comic Sans MS"
Text1.Font = "Comic Sans MS"
Else
End If
If Combo2.Text = "Fixedsys" Then
RichTextBox1.Font = "Fixedsys"
Text1.Font = "Fixedsys"
Else
End If
If Combo2.Text = "MS Serif" Then
RichTextBox1.Font = "MS Serif"
Text1.Font = "MS Serif"
Else
End If
End Sub

Private Sub Command1_Click()
RichTextBox1.SelBold = False
RichTextBox1.SelItalic = False
RichTextBox1.SelUnderline = False
RichTextBox1.SelStrikeThru = False
RichTextBox1.SelFontSize = 8
End Sub

Private Sub Command10_Click()
 Dim sfile As String
      On Error Resume Next
    With CommonDialog1
     .Filter = "All Files (*.rtf)|*.rtf"
    CommonDialog1.ShowSave
     sfile = .filename
     On Error Resume Next
    Open sfile For Output As 1
    Print #1, RichTextBox1.TextRTF
    End With
    Close
  Dim Counter As Integer
    Dim Workarea(25000) As String
    ProgressBar1.Min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.Min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter

Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.Min
  
End Sub

Private Sub Command11_Click()
  On Error Resume Next
 With CommonDialog1
        .Filter = "All Files (*.rtf)|*.rtf"
        .ShowOpen
        If Len(.filename) = 0 Then
        GoTo a
        On Error Resume Next
        End If
        sfar = .filename
        Kill sfar
End With
a:
End Sub

Private Sub Command12_Click()
 On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    
  On Error Resume Next
    With CommonDialog1
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        On Error Resume Next
        If ActiveForm.rtfText.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hDC
        End If
    On Error Resume Next
    End With

End Sub

Private Sub Command13_Click()
Dim asd As AlignmentConstants
Dim dfg As AlignmentConstants
asd = vbCenter
dfg = vbRightJustify
On Error Resume Next
RichTextBox1.RightMargin = dfg
RichTextBox1.SelAlignment = asd
End Sub

Private Sub Command14_Click()
Dim a As AlignmentConstants
Dim d As AlignmentConstants
a = vbLeftJustify
d = vbRightJustify
On Error Resume Next
RichTextBox1.RightMargin = d
RichTextBox1.SelAlignment = a
End Sub

Private Sub Command15_Click()
Dim s As AlignmentConstants
Dim f As AlignmentConstants
s = vbRightJustify
f = vbRightJustify
On Error Resume Next
RichTextBox1.RightMargin = f
RichTextBox1.SelAlignment = s
End Sub

Private Sub Command16_Click()
Dim am As String
On Error Resume Next
With CommonDialog1
        .Filter = "All Colors"
        .ShowColor
        On Error Resume Next
        am = .Color
RichTextBox1.SelColor = am
End With
End Sub

Private Sub Command17_Click()
Dim Search, Where
Search = InputBox("Enter text to be found:")
Where = InStr(RichTextBox1.Text, Search)
On Error Resume Next
If Where Then
RichTextBox1.SelStart = Where - 1
RichTextBox1.SelLength = Len(Search)
Else
MsgBox "String not found."
End If
End Sub



Private Sub Command18_Click()
Dim ch As String
Dim ur As String
ch = InputBox("Is Your Modem Switched ON (Yes/No)")
If ch = "Yes" Then
Load frmBrowser
frmBrowser.Visible = True
Dim u, p, a, b, c As String
u = InputBox("Enter Username", c)
p = InputBox("Enter Password", b)
Inet1.URL = ur
Inet1.Password = p
Else
End If
End Sub

Private Sub Command19_Click()
On Error Resume Next
With CommonDialog1
        .Filter = ("All Image Files (*.*)|*.*")
        .ShowOpen
    End With
    RichTextBox1.OLEObjects. _
    Add , , CommonDialog1.filename
End Sub

Private Sub Command2_Click()
If RichTextBox1.SelBold = True Then
RichTextBox1.SelBold = False
Else
RichTextBox1.SelBold = True
End If
End Sub


Private Sub Command20_Click()
Load Form1
Form1.Visible = True
End Sub

Private Sub Command21_Click()
Load frmEmployees
frmEmployees.Visible = True
End Sub



Private Sub Command23_Click()
Load Form2
Form2.Visible = True
End Sub

Private Sub Command25_Click()

End Sub

Private Sub Command24_Click()
Dim a As TabSelStyleConstants
a = tabTabStandard
End Sub

Private Sub Command3_Click()
If RichTextBox1.SelStrikeThru = True Then
RichTextBox1.SelStrikeThru = False
Else
RichTextBox1.SelStrikeThru = True
End If
End Sub

Private Sub Command4_Click()
Dim sfar As String
On Error Resume Next
    With CommonDialog1
        .Filter = "All Files (*.rtf)|*.rtf"
        .ShowOpen
        sfar = .filename
Dim Counter As Integer
    Dim Workarea(25000) As String
    ProgressBar1.Min = LBound(Workarea)
    ProgressBar1.Max = UBound(Workarea)
    ProgressBar1.Visible = True

'Set the Progress's Value to Min.
    ProgressBar1.Value = ProgressBar1.Min

'Loop through the array.
    For Counter = LBound(Workarea) To UBound(Workarea)
        'Set initial values for each item in the array.
        Workarea(Counter) = "Initial value" & Counter
        ProgressBar1.Value = Counter

Next Counter
    ProgressBar1.Visible = False
    ProgressBar1.Value = ProgressBar1.Min

Open sfar For Input As 1
RichTextBox1.TextRTF = Input$(LOF(1), 1)
End With
Close
End Sub

Private Sub Command5_Click()
If RichTextBox1.SelItalic = True Then
RichTextBox1.SelItalic = False
Else
RichTextBox1.SelItalic = True
End If
End Sub

Private Sub Command6_Click()
If RichTextBox1.SelUnderline = True Then
RichTextBox1.SelUnderline = False
Else
RichTextBox1.SelUnderline = True
End If
End Sub

Private Sub Command7_Click()
If RichTextBox1.SelBullet = True Then
RichTextBox1.SelBullet = False
Else
RichTextBox1.SelBullet = True
End If
End Sub

Private Sub Command8_Click()
Dim a As String
  With CommonDialog1
        .Filter = "all font"
        .ShowFont
        On Error Resume Next
        a = .Font
RichTextBox1.SelFontName = a
End With
End Sub

Private Sub Command9_Click()
Frame1.Visible = False
Combo2.Visible = False
List1.Visible = False
List2.Visible = False
Text1.Visible = False
End Sub

Private Sub COPY_Click()
MsgBox "Press Ctrl+C to copy the selected text"
End Sub

Private Sub CUT_Click()
MsgBox "Press Ctrl+X to cut the selected text"
End Sub

Private Sub PASTE_Clicks()
MsgBox "Press Ctrl+V to paste the copied/cutted text"
End Sub

Private Sub Form_Load()
Timer1.Interval = 1000
Combo2.Visible = False
Frame1.Visible = False
List2.Visible = False
List1.Visible = False
Text1.Visible = False
Dim name As String
name = InputBox("What is your Name Below:")
MsgBox "Hello, " & name & " Welcome to A++ Texteditor Ver 3.2 Pro"
List1.AddItem "Times"
List1.AddItem "MS"
List1.AddItem "Arial"
List1.AddItem "Bookman Old Style"
List1.AddItem "Century"
List1.AddItem "Arial Black"
List1.AddItem "Century Gothic"
List1.AddItem "Impact"
List1.AddItem "Bookman Antiqua"
List1.AddItem "Comic Sans MS"
List1.AddItem "Fixedsys"
List1.AddItem "MS Serif"
List2.AddItem "8"
List2.AddItem "10"
List2.AddItem "12"
List2.AddItem "14"
List2.AddItem "16"
List2.AddItem "18"
List2.AddItem "20"
List2.AddItem "22"
List2.AddItem "24"
List2.AddItem "26"
List2.AddItem "28"
List2.AddItem "30"
List2.AddItem "42"
List2.AddItem "72"
Slider1.Value = 0
End Sub

Private Sub List1_Click()
If List1.ListIndex = 0 Then
Combo2.Text = "Times New Roman"
Else
End If
If List1.ListIndex = 1 Then
Combo2.Text = "MS"
Else
End If
If List1.ListIndex = 2 Then
Combo2.Text = "Arial"
Else
End If
If List1.ListIndex = 3 Then
Combo2.Text = "Bookman Old Style"
Else
End If
If List1.ListIndex = 4 Then
Combo2.Text = "Century"
Else
End If
If List1.ListIndex = 5 Then
Combo2.Text = "Arial Black"
Else
End If
If List1.ListIndex = 6 Then
Combo2.Text = "Century Gothic"
Else
End If
If List1.ListIndex = 7 Then
Combo2.Text = "Impact"
Else
End If
If List1.ListIndex = 8 Then
Combo2.Text = "Bookman Antiqua"
Else
End If
If List1.ListIndex = 9 Then
Combo2.Text = "Comic Sans MS"
Else
End If
If List1.ListIndex = 10 Then
Combo2.Text = "Fixedsys"
Else
End If
End Sub

Private Sub List2_Click()
If List2.ListIndex = 0 Then
RichTextBox1.SelFontSize = 8
Text1.FontSize = 8
Else
End If
If List2.ListIndex = 1 Then
RichTextBox1.SelFontSize = 10
Text1.FontSize = 10
Else
End If
If List2.ListIndex = 2 Then
RichTextBox1.SelFontSize = 12
Text1.FontSize = 12
Else
End If
If List2.ListIndex = 3 Then
RichTextBox1.SelFontSize = 14
Text1.FontSize = 14
Else
End If
If List2.ListIndex = 4 Then
RichTextBox1.SelFontSize = 16
Text1.FontSize = 16
Else
End If
If List2.ListIndex = 5 Then
RichTextBox1.SelFontSize = 18
Text1.FontSize = 18
Else
End If
If List2.ListIndex = 6 Then
RichTextBox1.SelFontSize = 20
Text1.FontSize = 20
Else
End If
If List2.ListIndex = 7 Then
RichTextBox1.SelFontSize = 22
Text1.FontSize = 22
Else
End If
If List2.ListIndex = 8 Then
RichTextBox1.SelFontSize = 24
Text1.FontSize = 24
Else
End If
If List2.ListIndex = 9 Then
RichTextBox1.SelFontSize = 26
Text1.FontSize = 26
Else
End If
If List2.ListIndex = 10 Then
RichTextBox1.SelFontSize = 28
Text1.FontSize = 28
Else
End If
If List2.ListIndex = 11 Then
RichTextBox1.SelFontSize = 30
Text1.FontSize = 30
Else
End If
If List2.ListIndex = 12 Then
RichTextBox1.SelFontSize = 42
Text1.FontSize = 42
Else
End If
If List2.ListIndex = 13 Then
RichTextBox1.SelFontSize = 72
Text1.FontSize = 72
Else
End If
End Sub


Private Sub OLE2_Updated(Code As Integer)

End Sub

Private Sub Slider1_Click()
Dim a, x As AlignmentConstants
Dim d, Index As AlignmentConstants
If Slider1.Value = 1 Then
a = vbLeftJustify
d = vbRightJustify
x = (a + d) / 3
RichTextBox1.RightMargin = d
RichTextBox1.SelAlignment = x
Else
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 1000
Text2.Text = Time
End Sub
