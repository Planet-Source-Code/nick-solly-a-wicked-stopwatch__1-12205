VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStopWatch 
   BackColor       =   &H00000000&
   Caption         =   "Stop Watch"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   Icon            =   "frmStopWatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PBMin 
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      ToolTipText     =   "Minute Timer"
      Top             =   3000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   60
      Scrolling       =   1
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "Reset"
      Height          =   735
      Left            =   2640
      TabIndex        =   10
      ToolTipText     =   "Reset"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "Stop"
      Height          =   735
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Stop"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txt000 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      ToolTipText     =   "Counter"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtHours 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "Counter"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   ":"
      ToolTipText     =   "Counter"
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txt00 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      ToolTipText     =   "Counter"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtMins 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Counter"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txt0 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Counter"
      Top             =   240
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   1200
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Start"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   ":"
      ToolTipText     =   "Counter"
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtSecs 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "Counter"
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Minute Timer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "frmStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdReset_Click()
Timer1.Enabled = False
txtSecs = 0
txt0 = 0
txtMins = 0
txt00 = 0
txtHours = 0
txt000 = 0
PBMin.Value = 0
End Sub

Private Sub CmdStart_Click()
Timer1.Enabled = True

End Sub

Private Sub CmdStop_Click()
Timer1.Enabled = False

End Sub

Private Sub SoundRec1_GotFocus()

End Sub

Private Sub Timer1_Timer()
 If PBMin.Value = 60 Then
    PBMin.Value = 0
    
End If



    txtSecs = txtSecs + 1
   PBMin.Value = PBMin.Value + 1

If txtSecs = 10 Then
    txtSecs = 0
    txt0 = txt0 + 1
End If
If txt0 = 6 And txtSecs = 0 Then
    txtMins = txtMins + 1
    txtSecs = 0
    txt0 = 0
End If
If txtMins = 10 Then
    txtMins = 0
    txt00 = txt00 + 1
End If
If txt00 = 6 And txtMins = 0 Then
    txtHours = txtHours + 1
    txtMins = 0
    txt00 = 0
End If
If txtHours = 10 Then
    txtHours = 0
    txt000 = txt000 + 1
End If
If txt000 = 6 And txtHours = 0 Then
    Dim intpress As Integer
    intpress = MsgBox("End of Timer", vbOKOnly, Error)
    Timer1.Enabled = False
    txtSecs = 0
    txt0 = 0
    txtMins = 0
    txt00 = 0
    txtHours = 0
    txt000 = 0
End If
End Sub

Private Sub Timer2_Timer()
If txt100 = 10 Then
    txt100 = 0
    txt10 = txt10 + 1
End If
If txt100 = 0 And txt10 = 10 Then
    txt10 = 0
    txt100 = 0
End If
txt100 = txt100 + 1
End Sub
