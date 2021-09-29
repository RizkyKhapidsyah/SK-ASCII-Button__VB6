VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   351
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim previous As Integer
Private Sub Command1_Click(Index As Integer)
Command1(previous).BackColor = &H8000000F
Command1(previous).FontBold = False
Command1(Index).BackColor = &HFFFFFF
Command1(Index).FontBold = True
previous = Index
If Command1(Index).Caption <> "" Then
Form1.Caption = Command1(Index).Caption & ":" & Asc(Command1(Index).Caption) & ":&H" & Hex(Asc(Command1(Index).Caption))
Else
Form1.Caption = Command1(Index).Caption & ":" & 0 & ":&H0"
End If
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1(Index).Caption <> "" Then
Command1(Index).ToolTipText = Command1(Index).Caption & ":" & Asc(Command1(Index).Caption) & ":&H" & Hex(Asc(Command1(Index).Caption))
Else
Command1(Index).ToolTipText = Command1(Index).Caption & ":" & 0 & ":&H0"
End If

End Sub

Private Sub Form_Load()

previous = 0
Dim i As Integer
Form1.ScaleWidth = 16 * 26
Form1.ScaleHeight = 16 * 21
Show
Form1.Refresh
For X = 0 To 15
For Y = 0 To 15
 
CurrentX = X * 14
CurrentY = Y * 12
'Print Chr$(i)
Command1(i).Left = X * 26
Command1(i).Top = Y * 21
Command1(i).Width = 25
Command1(i).Height = 20
Command1(i).Caption = Chr$(i)


i = i + 1
Load Command1(i)

Next Y
Next X
For X = 0 To 255
Command1(X).Visible = True
Next
End Sub

