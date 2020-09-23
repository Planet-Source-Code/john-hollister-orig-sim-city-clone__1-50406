VERSION 5.00
Begin VB.Form Options 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SimTown Options"
   ClientHeight    =   1890
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Other Options..."
      Height          =   1695
      Left            =   1680
      TabIndex        =   10
      Top             =   120
      Width           =   2895
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Tax Rate %:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.OptionButton speed 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Game Speed"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton speed 
         Caption         =   "Option1"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton speed 
         Caption         =   "Option1"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton speed 
         Caption         =   "Option1"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Grueling"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Fast"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Medium"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Slow"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a As Integer
Option Explicit

Private Sub Form_Load()
form1.Enabled = False
form1.mo.Enabled = False
VScroll1.Value = 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
form1.Enabled = True
form1.mo.Enabled = True
taxRate = Text1.Text
End Sub

Private Sub Label1_Click(Index As Integer)
    For a = 1 To 4
        speed(a).Value = 0
    Next a
    speed(Index).Value = 1
End Sub

Private Sub OKButton_Click()

Unload Me
End Sub

Private Sub speed_Click(Index As Integer)
    For a = 1 To 4
        If a <> Index Then
            speed(a).Value = 0
        End If
    Next a

If Index = 1 Then
    form1.mo.Interval = 10000
End If
If Index = 2 Then
    form1.mo.Interval = 5000
End If
If Index = 3 Then
    form1.mo.Interval = 1000
End If
If Index = 4 Then
    form1.mo.Interval = 100
End If


End Sub


Private Sub VScroll1_Change()
  Text1.Text = (VScroll1.Value / 10)

End Sub
