VERSION 5.00
Begin VB.MDIForm TextEditor 
   BackColor       =   &H8000000C&
   Caption         =   "Text Editor"
   ClientHeight    =   3195
   ClientLeft      =   5550
   ClientTop       =   3645
   ClientWidth     =   4680
   Icon            =   "TextEditor.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu register 
      Caption         =   "&Register"
      Begin VB.Menu Click 
         Caption         =   "&Click"
      End
   End
End
Attribute VB_Name = "TextEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Click_Click()
Dim frm As New TextC
frm.Show
End Sub

Private Sub Exit_Click()
Unload Me
Set TextEditor = Nothing
End Sub

Private Sub New_Click()
Dim frm As New RichTextC
frm.Show
End Sub

Private Sub Open_Click()
Dim sm As String
With CommonDialog1
        .Filter = "All Files (*.rtf)|*.rtf"
        .ShowOpen
        sm = .filename
        On Error Resume Next
Open sm For Input As 1
RichTextBox2.TextRTF = Input$(LOF(1), 1)
End With
End Sub

Private Sub Save_as_Click()
 Dim sfile As String
    With dlgCommonDialog
     .Filter = "All Files (*.txt)|*.txt"
    dlgCommonDialog.ShowSave
     sfile = .filename
    End With
End Sub
