VERSION 5.00
Begin VB.Form TextC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   2640
   ClientLeft      =   5745
   ClientTop       =   3885
   ClientWidth     =   4035
   ControlBox      =   0   'False
   DrawMode        =   16  'Merge Pen
   Icon            =   "TextC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton okbutton 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Registration Key:"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Address:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Please fill in the details and get your Registration no. by mailing at :-)  aman_moudgil3000@yahoo.com"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "TextC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub okbutton_Click()
If Text3.Text = "111111" Then
Unload Me
Set TextC = Nothing
Load TextEditor
TextEditor.Visible = True
Else
MsgBox "Incorrect Key"
MsgBox "Your Application will Expire in 30 days"

Unload Me
Set TextC = Nothing
Load TextEditor
TextEditor.Visible = True
End If
End Sub
