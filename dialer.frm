VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dialer"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   3450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Get out of here"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "*"
      Height          =   375
      Index           =   10
      Left            =   1920
      TabIndex        =   12
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   2880
      TabIndex        =   11
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   2400
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   1920
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "#"
      Height          =   375
      Index           =   11
      Left            =   2880
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dial"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written by Mark Young 2000. Comments etc mark_a_young@worldnet.att.net

Private Sub Command1_Click()
Dim c As Integer
Dim n As String
Dim z As String

z = Combo1.Text
If z = "" Then Exit Sub
For c = 1 To Len(z)
    n = Mid$(z, c, 1)
    If IsNumeric(n) Then
        Call PlayWaveRes("T" & n)
        Command2(n).SetFocus
        Sleep 80
    Else
        Select Case n
            Case "#"
                Call PlayWaveRes("TPOUND")
                Command2(11).SetFocus
                Sleep 80
            Case "*"
                Call PlayWaveRes("TSTAR")
                Command2(10).SetFocus
                Sleep 80
        End Select
    End If
Me.Refresh
Next c
End Sub

Private Sub Command2_Click(Index As Integer)
Dim y As String

Select Case Index
    Case 10
        Combo1.Text = Combo1.Text & "*"
        Call PlayWaveRes("TSTAR")
    Case 11
        Combo1.Text = Combo1.Text & "#"
        Call PlayWaveRes("TPOUND")
    Case Else
        Combo1.Text = Combo1.Text & Index
        Call PlayWaveRes("T" & Index)
End Select

Me.Refresh
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Combo1.Text = ""
End Sub

Private Sub Form_Load()
Dim ff As Integer
Dim TempVar As String
Dim O As String

On Error GoTo errdesc

ff = FreeFile
O = App.Path & "\dialer.ini"
Open O For Input As ff
    Do While Not EOF(ff)
        Line Input #ff, TempVar
        Combo1.AddItem TempVar
    Loop
Close ff
Combo1.ListIndex = 0
Exit Sub
errdesc:
Select Case Err.Number
    Case 53
        s = MsgBox("File does not exist! Do you want to create it?", vbYesNo)
        If s = vbYes Then
            ff = FreeFile
            Open O For Output As ff
            Close ff
        End If
    Case Else
        MsgBox Err.Number & " " & Err.Description, vbCritical
End Select
End Sub
