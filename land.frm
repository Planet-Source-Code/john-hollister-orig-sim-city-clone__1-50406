VERSION 5.00
Begin VB.Form form1 
   ClientHeight    =   7320
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Views"
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   6480
      Width           =   7335
      Begin VB.CommandButton viewAllBut 
         Height          =   400
         Left            =   1800
         Picture         =   "land.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton viewIndBut 
         Height          =   400
         Left            =   1080
         Picture         =   "land.frx":06BA
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton viewComBut 
         Height          =   400
         Left            =   600
         Picture         =   "land.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton viewResBut 
         Height          =   400
         Left            =   120
         Picture         =   "land.frx":142E
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   400
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1280
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6075
      Width           =   5860
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5795
      Left            =   7150
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox Display 
      Height          =   375
      Left            =   7560
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame5 
      Caption         =   "Zoom"
      Height          =   735
      Left            =   7560
      TabIndex        =   17
      Top             =   5280
      Width           =   1095
      Begin VB.CommandButton zoomIn 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Picture         =   "land.frx":1AE8
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton zoomOut 
         Height          =   375
         Left            =   600
         Picture         =   "land.frx":21A2
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tools"
      Height          =   4575
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   1095
      Begin VB.Image stadBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   120
         Picture         =   "land.frx":285C
         Top             =   3600
         Width           =   405
      End
      Begin VB.Image demBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   120
         Picture         =   "land.frx":2F16
         Top             =   2160
         Width           =   405
      End
      Begin VB.Image roadBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   120
         Picture         =   "land.frx":35D0
         Top             =   240
         Width           =   405
      End
      Begin VB.Image ResBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   600
         Picture         =   "land.frx":3C8A
         Top             =   720
         Width           =   405
      End
      Begin VB.Image comBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   120
         Picture         =   "land.frx":4344
         Top             =   1200
         Width           =   405
      End
      Begin VB.Image indBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   600
         Picture         =   "land.frx":49FE
         Top             =   1200
         Width           =   405
      End
      Begin VB.Image powBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   120
         Picture         =   "land.frx":50B8
         Top             =   720
         Width           =   405
      End
      Begin VB.Image plaBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   120
         Picture         =   "land.frx":5772
         Top             =   3120
         Width           =   405
      End
      Begin VB.Image watBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   600
         Picture         =   "land.frx":5E2C
         Top             =   1680
         Width           =   405
      End
      Begin VB.Image policeBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   600
         Picture         =   "land.frx":64E6
         Top             =   2160
         Width           =   405
      End
      Begin VB.Image fireBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   120
         Picture         =   "land.frx":6BA0
         Top             =   2640
         Width           =   405
      End
      Begin VB.Image hospBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   600
         Picture         =   "land.frx":725A
         Top             =   2640
         Width           =   405
      End
      Begin VB.Image railBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   600
         Picture         =   "land.frx":7914
         Top             =   240
         Width           =   405
      End
      Begin VB.Image forBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   120
         Picture         =   "land.frx":7FCE
         Top             =   1680
         Width           =   405
      End
      Begin VB.Image airBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   600
         Picture         =   "land.frx":8688
         Top             =   3120
         Width           =   405
      End
      Begin VB.Image insBut 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   120
         Picture         =   "land.frx":8D42
         Top             =   4080
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info..."
      Height          =   5055
      Left            =   7560
      TabIndex        =   3
      Top             =   120
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Width           =   255
      End
      Begin VB.CommandButton unpause 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "land.frx":93FC
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton pause 
         Height          =   375
         Left            =   600
         Picture         =   "land.frx":9AB6
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.Frame Frame3 
         Caption         =   "Zoning:"
         Height          =   975
         Left            =   120
         TabIndex        =   12
         Top             =   3960
         Width           =   855
         Begin VB.Shape indbar 
            BackColor       =   &H0000FF00&
            BorderColor     =   &H00000000&
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   105
            Left            =   260
            Top             =   760
            Width           =   500
         End
         Begin VB.Shape combar 
            BackColor       =   &H0000FF00&
            BorderColor     =   &H00000000&
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   105
            Left            =   260
            Top             =   520
            Width           =   500
         End
         Begin VB.Shape resbar 
            BackColor       =   &H0000FF00&
            BorderColor     =   &H00000000&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   100
            Left            =   260
            Top             =   280
            Width           =   500
         End
         Begin VB.Label Label9 
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   75
            TabIndex        =   15
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Label2 
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   75
            TabIndex        =   14
            Top             =   480
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   80
            TabIndex        =   13
            Top             =   240
            Width           =   135
         End
      End
      Begin VB.Label label10 
         Caption         =   "Funds:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label money 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   3135
         Width           =   975
      End
      Begin VB.Label approv 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   2415
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Approval%:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Population:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label pop 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   1695
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label dateyear 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   495
         Width           =   975
      End
   End
   Begin VB.Frame currtool 
      Caption         =   "Current Tool"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      Begin VB.TextBox currprice 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.Image current 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   360
         Picture         =   "land.frx":A170
         Top             =   240
         Width           =   405
      End
   End
   Begin VB.PictureBox landframe 
      BackColor       =   &H00008080&
      Height          =   5655
      Left            =   1320
      ScaleHeight     =   5595
      ScaleWidth      =   5715
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   5775
      Begin VB.Image terrain 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1632
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Timer mo 
      Interval        =   1000
      Left            =   6960
      Top             =   6120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7560
      Top             =   6120
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   7440
      TabIndex        =   25
      Top             =   6600
      Width           =   975
   End
   Begin VB.Image ware 
      Height          =   345
      Index           =   3
      Left            =   8400
      Picture         =   "land.frx":A82A
      Top             =   6720
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image ware 
      Height          =   345
      Index           =   2
      Left            =   8760
      Picture         =   "land.frx":AEE4
      Top             =   6480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image ware 
      Height          =   345
      Index           =   1
      Left            =   9120
      Picture         =   "land.frx":B59E
      Top             =   6480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image bus 
      Height          =   345
      Index           =   3
      Left            =   8400
      Picture         =   "land.frx":BC58
      Top             =   6840
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image bus 
      Height          =   345
      Index           =   2
      Left            =   8760
      Picture         =   "land.frx":C312
      Top             =   6840
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image bus 
      Height          =   345
      Index           =   1
      Left            =   9120
      Picture         =   "land.frx":C9CC
      Top             =   6840
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image house 
      Height          =   345
      Index           =   3
      Left            =   8400
      Picture         =   "land.frx":D086
      Top             =   7200
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image house 
      Height          =   345
      Index           =   2
      Left            =   8760
      Picture         =   "land.frx":D740
      Top             =   7200
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image house 
      Height          =   345
      Index           =   1
      Left            =   9120
      Picture         =   "land.frx":DDFA
      Top             =   7200
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image noAir 
      Height          =   345
      Left            =   8760
      Picture         =   "land.frx":E4B4
      Top             =   6120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image plane 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":EB6E
      Top             =   6120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image air 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":F228
      Top             =   6120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image forest 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":F8E2
      Top             =   2520
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image nopowerail 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":FF9C
      Top             =   7560
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image powerail 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":10656
      Top             =   7560
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image railroad 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":10D10
      Top             =   1440
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image rail 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":113CA
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image bus 
      Height          =   345
      Index           =   0
      Left            =   9480
      Picture         =   "land.frx":11A84
      Top             =   6840
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image ware 
      Height          =   345
      Index           =   0
      Left            =   9480
      Picture         =   "land.frx":1213E
      Top             =   6480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image house 
      Height          =   345
      Index           =   0
      Left            =   9480
      Picture         =   "land.frx":127F8
      Top             =   7200
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image noHosp 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":12EB2
      Top             =   5760
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image noFire 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":1356C
      Top             =   5400
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image noPolice 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":13C26
      Top             =   5040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image hospital 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":142E0
      Top             =   5760
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image fire 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":1499A
      Top             =   5400
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image police 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":15054
      Top             =   5040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image noPoweroad 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":1570E
      Top             =   4320
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image noInd 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":15DC8
      Top             =   3600
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image noCom 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":16482
      Top             =   3240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image noRes 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":16B3C
      Top             =   2880
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image nopower 
      Height          =   345
      Left            =   9120
      Picture         =   "land.frx":171F6
      Top             =   3960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image plant 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":178B0
      Top             =   4680
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image poweroad 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":17F6A
      Top             =   4320
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image power 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":18624
      Top             =   3960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image ind 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":18CDE
      Top             =   3600
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image com 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":19398
      Top             =   3240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image res 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":19A52
      Top             =   2880
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image activeWater 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":1A10C
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image activeLand 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":1A7C6
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image land 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":1AE80
      Top             =   1800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image active 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":1B53A
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image road 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":1BBF4
      Top             =   960
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image water 
      Height          =   345
      Left            =   9480
      Picture         =   "land.frx":1C2AE
      Top             =   360
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Menu optionMenu 
      Caption         =   "Options"
      Begin VB.Menu opt_menu 
         Caption         =   "Options"
      End
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xlast, ylast As Single
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim p(200, 200)
Private Const VK_SPACE = &H20
Private Const VK_DOWN = &H28
Private Const VK_doubleu = &H57
Private Const VK_SPACE1 = &H31
Private Const VK_ay = &H41
Private Const VK_dee = &H44
Private Const VK_ex = &H58
Private Const VK_UP = &H26
Private Const VK_LEFT = &H25
Private Const VK_RIGHT = &H27
Public lay As Integer
Public month As Integer
Public numPlants As Integer
Public currentPrice As Integer
Public lastV, lastH As Integer
Dim numFire, numPolice, numHosp As Integer


Private Sub airBut_Click()
lay = 16
currprice.Text = 100
current.Picture = airBut.Picture
    currtool.Caption = "Airport"
End Sub

Private Sub comBut_Click()
lay = 4 '4 = commercial
currprice.Text = 25
current.Picture = comBut.Picture
    currtool.Caption = "Commercial"
End Sub





Private Sub command5_Click()

Load Report
Report.Show

End Sub



Private Sub demBut_Click()
    lay = -4
    current.Picture = demBut.Picture
    currprice.Text = 15
    currtool.Caption = "Demolish"
End Sub

Private Sub fireBut_Click()
lay = 9
currprice.Text = 500
current.Picture = fireBut.Picture
    currtool.Caption = "Fire"
End Sub

Private Sub forBut_Click()
lay = -2
currprice.Text = 20
current.Picture = forBut.Picture
    currtool.Caption = "Forest/Park"
End Sub

Private Sub Form_Load()
Call newGame

VScroll1.Max = 24
HScroll1.Max = 24

End Sub

Function newGame()
numPlants = 0
lay = -3
month = 0
taxRate = 5
funds = 4000
population = 0
jobs = 0
year = 1940
lastV = 0
lastH = 0
numFire = 0
numHosp = 0
numPolice = 0


monArray(1) = "Jan, "
monArray(2) = "Feb, "
monArray(3) = "Mar, "
monArray(4) = "Apr, "
monArray(5) = "May, "
monArray(6) = "Jun, "
monArray(7) = "Jul, "
monArray(8) = "Aug, "
monArray(9) = "Sep, "
monArray(10) = "Oct, "
monArray(11) = "Nov, "
monArray(12) = "Dec, "



condvals(0) = 0
condvals(1) = 1
condvals(9) = 2
condvals(3) = -1
condvals(10) = -2
condvals(5) = -40
condvals(6) = -41
condvals(7) = -39
condvals(8) = 40
condvals(2) = 41
condvals(4) = 39
condvals(11) = -80
condvals(12) = 80
Randomize




    For a = 0 To 39
        For b = 0 To 39
            tmp = (a * 40) + b
            Load terrain(tmp)
            terrain(tmp).Left = b * 360
            terrain(tmp).Top = a * 360
            terrain(tmp).Picture = land.Picture
            powergrid(tmp) = 0
            propVal(RESID, tmp) = 10
            propVal(COMME, tmp) = 10
            propVal(INDUS, tmp) = 10
        Next b
    Next a
    For a = 0 To 200
        tmfor = Int(Rnd * 1600)
        terrArr(tmfor) = -2
        terrain(tmfor) = forest.Picture
    Next a
    landframe.Width = (terrain(0).Width - 10) * 16
    landframe.Height = (terrain(0).Height - 10) * 16
    Call makeFrame
    
    'make river
    Randomize
    lastspot = Int(Rnd * 25) + 5
    For a = 0 To 39
        tmp = Int((Rnd * 3) - 1)
        If lastspot + tmp < 0 Then
            tmp = 0
        End If
        If lastspot + tmp >= 36 Then
            tmp = 0
        End If
        For b = 0 To 2
            terrain(a * 40 + (lastspot + tmp + b)).Picture = water.Picture
            terrArr(a * 40 + (lastspot + tmp + b)) = 1
        Next b
        lastspot = (lastspot + tmp)
    Next a

End Function




Private Sub hospBut_Click()
lay = 10
currprice.Text = 500
current.Picture = hospBut.Picture
    currtool.Caption = "Hospital"
End Sub

Private Sub HScroll1_Change()
If zoomOut.Enabled = True Then
    If HScroll1.Value > lastH Then
        Call invisible
        For lastH = lastH To HScroll1.Value - 1
            For a = 0 To 39
                For b = 0 To 39
                    tmp = (a * 40) + b
                    terrain(tmp).Left = terrain(tmp).Left - 360
                Next b
            Next a
        Next lastH
        Call revisible
    Else
        Call invisible
        For lastH = lastH To HScroll1.Value + 1 Step -1
            For a = 0 To 39
                For b = 0 To 39
                    tmp = (a * 40) + b
                    terrain(tmp).Left = terrain(tmp).Left + 360
                Next b
            Next a
        Next lastH
        Call revisible
   End If
End If
lastH = HScroll1.Value
Call makeFrame

End Sub

Private Sub insBut_Click()
    lay = -3
currprice.Text = 0
current.Picture = insBut.Picture
    currtool.Caption = "Inspect"
End Sub

Private Sub opt_menu_Click()
    Load Options
    Options.Show
End Sub

Private Sub pause_Click()
    mo.Enabled = False
    pause.Enabled = False
    unpause.Enabled = True
End Sub

Private Sub railBut_Click()
    lay = -1
    currprice.Text = 20
    current.Picture = railBut.Picture
    currtool.Caption = "Railroad"
End Sub

Private Sub stadBut_Click()
    lay = 17
    currprice.Text = 1500
    current.Picture = stadBut.Picture
    If population > 700 Then
        currtool.Caption = "Stadium"
    Else
        currtool.Caption = "Unavailable"
    End If
End Sub

Private Sub unpause_Click()
    mo.Enabled = True
    unpause.Enabled = False
    pause.Enabled = True
End Sub

Private Sub viewAllBut_Click()
For a = 0 To 1599
    terrain(a).Visible = True
Next a
End Sub

Private Sub viewComBut_Click()
For a = 0 To 1599
    terrain(a).Visible = False
    If terrArr(a) = 14 Or terrArr(a) = 4 Or terrArr(a) = 1 Then
        terrain(a).Visible = True
    End If
Next a
End Sub

Private Sub viewIndBut_Click()
For a = 0 To 1599
    terrain(a).Visible = False
    If terrArr(a) = 15 Or terrArr(a) = 5 Or terrArr(a) = 7 Or terrArr(a) = 1 Then
        terrain(a).Visible = True
    End If
Next a
End Sub

Private Sub viewResBut_Click()
For a = 0 To 1599
    terrain(a).Visible = False
    If terrArr(a) = 13 Or terrArr(a) = 3 Or terrArr(a) = 1 Then
        terrain(a).Visible = True
    End If
Next a
End Sub

Private Sub VScroll1_Change()
If zoomOut.Enabled = True Then
    If VScroll1.Value > lastV Then
        Call invisible
        For lastV = lastV To VScroll1.Value - 1
            For a = 0 To 39
                For b = 0 To 39
                    tmp = (a * 40) + b
                    terrain(tmp).Top = terrain(tmp).Top - 360
                Next b
            Next a
        Next lastV
        Call revisible
    Else
        Call invisible
        For lastV = lastV To VScroll1.Value + 1 Step -1
            For a = 0 To 39
                For b = 0 To 39
                    tmp = (a * 40) + b
                    terrain(tmp).Top = terrain(tmp).Top + 360
                Next b
            Next a
        Next lastV
        Call revisible
   End If
End If
lastV = VScroll1.Value
Call makeFrame
End Sub

Private Sub watBut_Click()
 lay = 1  '1 = water
currprice.Text = 20
current.Picture = watBut.Picture
    currtool.Caption = "Water"
End Sub
Private Sub indBut_Click()
    lay = 5 '5 = industrial
    currprice.Text = 25
    current.Picture = indBut.Picture
    currtool.Caption = "Industrial"
End Sub

Private Sub Label2_Click()
Call newGame
End Sub

Private Sub plaBut_Click()
    lay = 7 '7 = powerplant
currprice.Text = 1000
current.Picture = plaBut.Picture
    currtool.Caption = "Power Plant"
End Sub

Private Sub policeBut_Click()
    lay = 8
currprice.Text = 500
current.Picture = policeBut.Picture
    currtool.Caption = "Police"
End Sub

Private Sub powBut_Click()
    lay = 6  '6 = powerline
    currprice.Text = 20
    current.Picture = powBut.Picture
    currtool.Caption = "Power Lines"
End Sub

Private Sub ResBut_Click()
    lay = 3  '3 = residential
    currprice.Text = 25
    current.Picture = ResBut.Picture
    currtool.Caption = "Residential"
End Sub

Private Sub roadBut_Click()
    lay = 2 '2 = road
    currprice.Text = 20
    current.Picture = roadBut.Picture
    currtool.Caption = "Road"
End Sub

Private Sub terrain_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If X <= terrain(Index).Width And Y <= terrain(Index).Height And X >= 0 And Y >= 0 And funds >= currprice.Text Then
    If lay = -3 And Button = 1 Then
        display.Picture = terrain(Index).Picture
        sendRep = Index
        Load Report
        Report.Show
    End If
    
    If lay = -2 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Then
            terrain(Index).Picture = forest.Picture
            funds = funds - 20
            terrArr(Index) = lay
        End If
    End If
    
    If lay = -1 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Or terrain(Index).Picture = activeWater.Picture Then
            terrain(Index).Picture = rail.Picture
            funds = funds - 20
            terrArr(Index) = lay
        End If
        If terrain(Index).Picture = road.Picture Then
            terrain(Index).Picture = railroad.Picture
            funds = funds - 20
            terrArr(Index) = lay
        End If
        If terrain(Index).Picture = power.Picture Then
            terrain(Index).Picture = powerail.Picture
            funds = funds - 20
            terrArr(Index) = 22
        End If
    End If
    
    If lay = 1 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Then
            terrain(Index).Picture = water.Picture
            funds = funds - 20
            terrArr(Index) = lay
        End If
    End If
    
    If lay = 2 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Or terrain(Index).Picture = activeWater.Picture Then
            terrain(Index).Picture = road.Picture
            funds = funds - 20
            terrArr(Index) = lay
        End If
        If terrain(Index).Picture = power.Picture Then
            terrain(Index).Picture = poweroad.Picture
            funds = funds - 20
            terrArr(Index) = 12
        End If
        If terrain(Index).Picture = rail.Picture Then
            terrain(Index).Picture = railroad.Picture
            funds = funds - 20
            terrArr(Index) = lay
        End If
    End If
    
    If lay = 3 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Then
            funds = funds - 25
            terrain(Index).Picture = res.Picture
            terrArr(Index) = lay
            propertyVal (Index)
       End If
    End If
    
     If lay = 4 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Then
            funds = funds - 25
            terrain(Index).Picture = com.Picture
            terrArr(Index) = lay
            propertyVal (Index)
        End If
    End If
    
    If lay = 5 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Then
            funds = funds - 25
            terrain(Index).Picture = ind.Picture
            terrArr(Index) = lay
            propertyVal (Index)
        End If
    End If
    
    If lay = 6 And Button = 1 Then
        If terrain(Index).Picture = road.Picture Then
            funds = funds - 20
            terrain(Index).Picture = poweroad.Picture
            lay = 12
            terrArr(Index) = lay
        End If
        If terrain(Index).Picture = rail.Picture Then
            funds = funds - 20
            terrain(Index).Picture = powerail.Picture
            lay = 22
            terrArr(Index) = lay
        End If
        
        If terrain(Index).Picture = activeLand.Picture Or terrain(Index).Picture = activeWater.Picture Then
            funds = funds - 20
            terrain(Index).Picture = power.Picture
            terrArr(Index) = lay
        End If
    End If
    
     If lay = 7 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Then
            funds = funds - 1000
            terrain(Index).Picture = plant.Picture
            jobs = jobs + 15
            powergrid(Index) = 2
            numPlants = numPlants + 1
            terrArr(Index) = lay
            Call pollute(Index, 5)
            If (Index - 41) >= 0 Then
                Call pollute(Index - 41, 5)
            End If
            If (Index - 40) >= 0 Then
                Call pollute(Index - 40, 5)
            End If
            If (Index - 39) >= 0 Then
                Call pollute(Index - 39, 5)
            End If
            If (Index - 1) >= 0 Then
                Call pollute(Index - 1, 5)
            End If
            If (Index + 1) <= 1599 Then
                Call pollute(Index + 1, 5)
            End If
            If (Index + 41) <= 1599 Then
                Call pollute(Index + 41, 5)
            End If
            If (Index + 40) <= 1599 Then
                Call pollute(Index + 40, 5)
            End If
            If (Index + 39) <= 1599 Then
                Call pollute(Index + 39, 5)
            End If
      End If
    End If
    
    If lay = 8 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Then
            funds = funds - 500
            numPolice = numPolice + 1
            terrain(Index).Picture = police.Picture
            jobs = jobs + 10
            terrArr(Index) = lay
            Call makeCrime(Index, -3)
            If (Index - 41) >= 0 Then
                Call makeCrime(Index - 41, -3)
            End If
            If (Index - 40) >= 0 Then
                Call makeCrime(Index - 40, -3)
            End If
            If (Index - 39) >= 0 Then
                Call makeCrime(Index - 39, -3)
            End If
            If (Index - 1) >= 0 Then
                Call makeCrime(Index - 1, -3)
            End If
            If (Index + 1) <= 1599 Then
                Call makeCrime(Index + 1, -3)
            End If
            If (Index + 41) <= 1599 Then
                Call makeCrime(Index + 41, -3)
            End If
            If (Index + 40) <= 1599 Then
                Call makeCrime(Index + 40, -3)
            End If
            If (Index + 39) <= 1599 Then
                Call makeCrime(Index + 39, -3)
            End If
        End If
    End If
    
    If lay = 9 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Then
            funds = funds - 500
            terrain(Index).Picture = fire.Picture
            jobs = jobs + 10
            numFire = numFire + 1
            terrArr(Index) = lay

            Call makeFire(Index, -3)
            If (Index - 41) >= 0 Then
                Call makeFire(Index - 41, -3)
            End If
            If (Index - 40) >= 0 Then
                Call makeFire(Index - 40, -3)
            End If
            If (Index - 39) >= 0 Then
                Call makeFire(Index - 39, -3)
            End If
            If (Index - 1) >= 0 Then
                Call makeFire(Index - 1, -3)
            End If
            If (Index + 1) <= 1599 Then
                Call makeFire(Index + 1, -3)
            End If
            If (Index + 41) <= 1599 Then
                Call makeFire(Index + 41, -3)
            End If
            If (Index + 40) <= 1599 Then
                Call makeFire(Index + 40, -3)
            End If
            If (Index + 39) <= 1599 Then
                Call makeFire(Index + 39, -3)
            End If
        End If
    End If
    
    If lay = 10 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Then
            funds = funds - 500
            terrain(Index).Picture = hospital.Picture
            jobs = jobs + 10
            terrArr(Index) = lay
                Call makecare(Index, 2)
            If (Index - 41) >= 0 Then
                Call makecare(Index - 41, 2)
            End If
            If (Index - 40) >= 0 Then
                Call makecare(Index - 40, 2)
            End If
            If (Index - 39) >= 0 Then
                Call makecare(Index - 29, 2)
            End If
            If (Index - 1) >= 0 Then
                Call makecare(Index - 1, 2)
            End If
            If (Index + 1) <= 1599 Then
                Call makecare(Index + 1, 2)
            End If
            If (Index + 41) <= 1599 Then
                Call makecare(Index + 41, 2)
            End If
            If (Index + 40) <= 1599 Then
                Call makecare(Index + 40, 2)
            End If
            If (Index + 39) <= 1599 Then
                Call makecare(Index + 39, 2)
            End If

        End If
    End If
    
    
    If lay = 16 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Then
            funds = funds - 100
            terrain(Index).Picture = air.Picture
            terrArr(Index) = lay
       End If
    End If
    
     If lay = 17 And Button = 1 Then
        If terrain(Index).Picture = activeLand.Picture Then
            If population > 700 Then
                funds = funds - 1500
                terrain(Index).Picture = stadBut.Picture
                terrArr(Index) = lay
            End If
       End If
    End If
    

    
    If lay = 12 Or lay = 22 Then
        lay = 6
    End If
  
'--- bulldoze  -----
  
    If Button = 2 Or lay = -4 Then
        powergrid(Index) = 0
        If terrArr(Index) = 8 Then
            numPolice = numPolice - 1
            Call makeCrime(Index, 3)
            If (Index - 41) >= 0 Then
                Call makeCrime(Index - 41, 3)
            End If
            If (Index - 40) >= 0 Then
                Call makeCrime(Index - 40, 3)
            End If
            If (Index - 39) >= 0 Then
                Call makeCrime(Index - 39, 3)
            End If
            If (Index - 1) >= 0 Then
                Call makeCrime(Index - 1, 3)
            End If
            If (Index + 1) <= 1599 Then
                Call makeCrime(Index + 1, 3)
            End If
            If (Index + 41) <= 1599 Then
                Call makeCrime(Index + 41, 3)
            End If
            If (Index + 40) <= 1599 Then
                Call makeCrime(Index + 40, 3)
            End If
            If (Index + 39) <= 1599 Then
                Call makeCrime(Index + 39, 3)
            End If
        End If
        
        If terrArr(Index) = 9 Then
            numFire = numFire - 1
            Call makeFire(Index, 3)
            If (Index - 41) >= 0 Then
                Call makeFire(Index - 41, 3)
            End If
            If (Index - 40) >= 0 Then
                Call makeFire(Index - 40, 3)
            End If
            If (Index - 39) >= 0 Then
                Call makeFire(Index - 39, 3)
            End If
            If (Index - 1) >= 0 Then
                Call makeFire(Index - 1, 3)
            End If
            If (Index + 1) <= 1599 Then
                Call makeFire(Index + 1, 3)
            End If
            If (Index + 41) <= 1599 Then
                Call makeFire(Index + 41, 3)
            End If
            If (Index + 40) <= 1599 Then
                Call makeFire(Index + 40, 3)
            End If
            If (Index + 39) <= 1599 Then
                Call makeFire(Index + 39, 3)
            End If
        End If
        
         If terrArr(Index) = 9 Then
            numHosp = numHosp - 1
            Call makecare(Index, 2)
            If (Index - 41) >= 0 Then
                Call makecare(Index - 41, -2)
            End If
            If (Index - 40) >= 0 Then
                Call makecare(Index - 40, -2)
            End If
            If (Index - 39) >= 0 Then
                Call makecare(Index - 39, -2)
            End If
            If (Index - 1) >= 0 Then
                Call makecare(Index - 1, -2)
            End If
            If (Index + 1) <= 1599 Then
                Call makecare(Index + 1, -2)
            End If
            If (Index + 41) <= 1599 Then
                Call makecare(Index + 41, -2)
            End If
            If (Index + 40) <= 1599 Then
                Call makecare(Index + 40, -2)
            End If
            If (Index + 39) <= 1599 Then
                Call makecare(Index + 39, -2)
            End If
        End If
       
        
        If terrArr(Index) = 14 Then
            jobs = jobs - comJob(Index)
            comJob(Index) = 0
            comtype(Index) = 0
            Call makeCrime(Index, -1)
        End If
        If terrArr(Index) = 13 Then
            population = population - occupant(Index)
            occupant(Index) = 0
            housetype(Index) = 0
            Call makeFire(Index, -1)
        End If
        If terrArr(Index) = 15 Then
            jobs = jobs - indjob(Index)
            indjob(Index) = 0
            indtype(Index) = 0
            Call pollute(Index, -1)
        End If
        If powergrid(Index) = 2 Then
            jobs = jobs - 15
        End If
        For a = 0 To 1599
            If powergrid(a) <> 2 Then
                powergrid(a) = 0
            End If
        Next a
        funds = funds - 15
        terrain(Index).Picture = land.Picture
        terrArr(Index) = 0
        Call havePower
    End If
End If
End Sub

Private Sub terrain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For a = 0 To 1599
        If terrain(a).Visible = True Then
            If terrain(a).Picture = activeWater.Picture Then
                terrain(a).Picture = water.Picture
            End If
            If terrain(a).Picture = activeLand.Picture Then
                terrain(a).Picture = land.Picture
            End If
        End If
    Next a
    If terrain(Index).Picture = land.Picture Then
        terrain(Index).Picture = activeLand.Picture
    End If
    If terrain(Index).Picture = water.Picture Then
        terrain(Index).Picture = activeWater.Picture
    End If
   
   
End Sub

Private Sub Timer1_Timer()
Dim tmpoll As Integer
Label3.Caption = jobs
For a = 0 To 1599
     tmpoll = tmpoll + pollution(a)
Next a

money.Caption = "$" & funds
End Sub


Function makeFrame()
    For a = 0 To 1599
        If terrain(a).Top >= 0 And terrain(a).Top < landframe.Height Then
            If terrain(a).Left >= 0 And terrain(a).Left < landframe.Width Then
                terrain(a).Visible = True
            End If
        End If
    Next a
End Function


Function invisible()
    landframe.Visible = False
        For a = 0 To 39
            For b = 0 To 39
                tmp = (a * 40) + b
                terrain(tmp).Visible = False
            Next b
        Next a

End Function

Function revisible()
    landframe.Visible = True
        For a = 0 To 39
            For b = 0 To 39
                tmp = (a * 40) + b
                terrain(tmp).Visible = True
            Next b
        Next a
End Function

Private Sub mo_Timer()
Timer1.Enabled = False
    month = (month Mod 12) + 1

If month = 1 Then
    year = year + 1
    Call calcApproval
    'put REVENUE HERE--------------
    
    Dim tmprev As Double
    temprev = (approval / 100) * (population * (3 + (taxRate / 100)) - numPlants * 10)
    temprev = temprev - (0.5 * population * numPolice)
    temprev = temprev - (0.5 * population * numHosp)
    temprev = temprev - (0.5 * population * numFire)
    
    funds = Int(temprev) + funds
    If funds > 99999 Then
        funds = 99999
    End If
    approv.Caption = approval
End If

'------info stats update---------
dateyear.Caption = monArray(month) & year
pop.Caption = population




Dim a As Integer
    
    totalRes = 0
    totalCom = 0
    totalInd = 0
    totalAir = 0
    For a = 0 To 1599
        If terrArr(a) = 13 Then
            totalRes = totalRes + 1
        End If
        If terrArr(a) = 14 Then
            totalCom = totalCom + 1
        End If
        If terrArr(a) = 15 Then
            totalInd = totalInd + 1
        End If
        If terrArr(a) = 7 Then
            totalInd = totalInd + 1
        End If
        If terrArr(a) = 26 Then
            totalAir = totalAir + 1
        End If
    Next a
        
    
    'extensive testing!!!!!!
    Call havePower
    
    
'NEED TO ADD A GROWTH RATE!
'ADD:
'TAXES


Dim tmp As Integer
Dim moveout As Integer
Dim movein As Integer
Dim justleft As Boolean
justleft = False


'airports-----------------------

For a = 0 To 1599
    If terrArr(a) = 26 Then
        If powergrid(a) = 0 Or population < 500 Then
            If terrain(a).Picture = plane.Picture Then
                jobs = jobs - 6
                Call makenoise(a, -1)
            End If
            terrain(a).Picture = air.Picture
            terrArr(a) = 16
        End If
    End If


    If terrArr(a) = 16 And jobs > 0 Then
        If powergrid(a) = 1 And population > 500 And (totalRes - 150) / (totalAir + 1) > 10 Then
        If (totalInd - 10) / (totalAir + 1) > 5 And (totalCom - 10) / (totalAir + 1) > 4 Then
            terrain(a).Picture = plane.Picture
            If terrArr(a) = 5 Then
                Call makenoise(a, 1)
            End If
            totalAir = totalAir + 1
            terrArr(a) = 26
            jobs = jobs + 6
        End If
        End If
    End If
    


    Call propertyVal(a)




'----INDUSTRY-----------------


    If terrArr(a) = 15 Then

        Randomize
        moveout = Int(Rnd * 100)
        If powergrid(a) = 0 Then
            If terrArr(a) = 15 Then
                jobs = jobs - indjob(a)
                Call pollute(a, -1)
            End If
            terrain(a).Picture = ind.Picture
            terrArr(a) = 5
            indjob(a) = 0
            indtype(a) = 0
        End If

        If moveout = 0 Then
            If terrArr(a) = 15 Then
                jobs = jobs - indjob(a)
                Call pollute(a, -1)
            End If
            terrain(a).Picture = ind.Picture
            terrArr(a) = 5
            justleft = True
            indjob(a) = 0
            indtype(a) = 0
        End If
    End If


    If terrArr(a) = 5 And justleft = False And jobs > 0 Then

        Randomize
        movein = Int(Rnd * (36 - (propVal(INDUS, a) / 4)))
        If movein = 0 And powergrid(a) = 1 And Int(population / jobs) > 0 Then
            Randomize
            tmp = Int(Rnd * 2) + 1
            tmpjob = Int(Rnd * 4) + 1
            If propVal(INDUS, a) <= 15 And indtype(a) = 0 Then
                indtype(a) = tmp
                terrain(a).Picture = ware(indtype(a) - 1).Picture
                indjob(a) = tmpjob
                jobs = jobs + tmpjob
            End If
            If propVal(INDUS, a) > 15 And indtype(a) = 0 Then
                indtype(a) = tmp + 2
                terrain(a).Picture = ware(indtype(a) - 1).Picture
                indjob(a) = tmpjob * 3
                jobs = jobs + tmpjob * 3
            End If
            If terrArr(a) = 15 And terrain(a).Picture = ind.Picture Then
                terrain(a).Picture = ware(indtype(a) - 1).Picture
                indjob(a) = 15
                jobs = jobs + 15
            End If
            terrArr(a) = 15
            Call pollute(a, 1)
        End If
    End If
    If terrArr(a) = 15 And terrain(a).Picture = ind.Picture Then
        terrain(a).Picture = ware(indtype(a) - 1).Picture
    End If
'------COMMERCE----------------------
    
    If terrArr(a) = 14 Then

        Randomize
        moveout = Int(Rnd * 100)
        If powergrid(a) = 0 Then
            If terrArr(a) = 14 Then
                jobs = jobs - comJob(a)
                Call makeCrime(a, -1)
            End If
            terrain(a).Picture = com.Picture
            terrArr(a) = 4
            comJob(a) = 0
            comtype(a) = 0
        End If

        If moveout = 0 Then
            If terrArr(a) = 14 Then
                jobs = jobs - comJob(a)
            Call makeCrime(a, -1)
            End If
            terrain(a).Picture = com.Picture
            terrArr(a) = 4
            justleft = True
            comJob(a) = 0
            comtype(a) = 0
        End If
    End If


    If terrArr(a) = 4 And justleft = False And jobs > 0 Then

        Randomize
        movein = Int(Rnd * (36 - (propVal(COMME, a) / 4)))
        If movein = 0 And powergrid(a) = 1 And Int(population / jobs) > 0 Then
            Randomize
            tmp = Int(Rnd * 2) + 1
            tmpjob = Int(Rnd * 4) + 1
            If propVal(COMME, a) <= 15 And comtype(a) = 0 Then
                comtype(a) = tmp
                terrain(a).Picture = bus(comtype(a) - 1).Picture
                comJob(a) = tmpjob
                jobs = jobs + comJob(a)
            End If
            If propVal(COMME, a) > 15 And comtype(a) = 0 Then
                comtype(a) = tmp + 2
                terrain(a).Picture = bus(comtype(a) - 1).Picture
                comJob(a) = tmpjob * 3
                jobs = jobs + comJob(a)
            End If
            If terrArr(a) = 14 And terrain(a).Picture = com.Picture Then
                terrain(a).Picture = bus(comtype(a) - 1).Picture
                comJob(a) = comtype(a) * 2
                jobs = jobs + comtype(a) * 2
            End If
            terrArr(a) = 14
            Call makeCrime(a, 1)
        End If
    End If
    If terrArr(a) = 14 And terrain(a).Picture = com.Picture Then
        terrain(a).Picture = bus(comtype(a) - 1).Picture
    End If

'---HOUSING------------------------

justleft = False
    If terrArr(a) = 13 Then

        Randomize
        moveout = Int(Rnd * 100)
        If propVal(RESID, a) < 0 And powergrid(a) = 1 Then
            If terrArr(a) = 13 Then
                population = population - occupant(a)
                occupant(a) = 0
            End If
            terrain(a).Picture = res.Picture
            terrArr(a) = 3
            Call makeFire(a, -1)
        End If
 
        If moveout = 0 Or powergrid(a) = 0 Then
            If terrArr(a) = 13 Then
                population = population - occupant(a)
                occupant(a) = 0
            End If
            terrain(a).Picture = res.Picture
            terrArr(a) = 3
            justleft = True
            Call makeFire(a, -1)
        End If

    End If

    If terrArr(a) = 3 And justleft = False Then

        movein = Int(Rnd * (30 - (propVal(RESID, a) / 4)))
        If (propVal(RESID, a) > 0 Or movein = 0) And powergrid(a) = 1 Then
        
        
        'magic NUMBER!!!!!!!!!!!!!!!!!--------------- 0.8
                               '|        > 1 population falls
                               'V        < 1 population grows
                                        'depend on approval rating!
        If approval <= 0 Then
            approval = 1
        End If
        If jobs > (population) * (0.8) Then
            terrArr(a) = 13
            Randomize
            tmp = Int(Rnd * 2) + 1
            If propVal(RESID, a) <= 15 And housetype(a) = 0 Then
                housetype(a) = tmp
                terrain(a).Picture = house(housetype(a) - 1).Picture
            End If
            If propVal(RESID, a) > 15 And housetype(a) = 0 Then
                housetype(a) = tmp + 2
                terrain(a).Picture = house(housetype(a) - 1).Picture
            End If
            If terrArr(a) = 13 And terrain(a).Picture = res.Picture Then
                terrain(a).Picture = house(housetype(a) - 1).Picture
            End If
            Randomize
            peop = Int(Rnd * 5) + 1
            occupant(a) = peop
            population = population + peop
            Call makeFire(a, 1)
        End If
        End If
       
    End If
Next a

Timer1.Enabled = True

End Sub


Function pollute(loc As Integer, c As Integer)
'c decides whether crime is made or cleaned
    If powergrid(loc) > 0 Then
        pollution(loc) = pollution(loc) + (c * 3)
        If (loc - 1) Mod 40 > 0 And loc - 2 >= 0 Then
            pollution(loc - 2) = pollution(loc - 2) + (c * 1)
        End If
        If (loc) Mod 40 > 0 And loc - 1 >= 0 And loc - 41 >= 0 And loc + 39 <= 1599 Then
            pollution(loc - 1) = pollution(loc - 1) + (c * 2)
            pollution(loc - 41) = pollution(loc - 41) + (c * 1)
            pollution(loc + 39) = pollution(loc + 39) + (c * 1)
        End If
        If (loc) Mod 40 < 39 And loc + 1 <= 1599 And loc + 41 <= 1599 And loc - 39 >= 0 Then
            pollution(loc + 1) = pollution(loc + 1) + (c * 2)
            pollution(loc + 41) = pollution(loc + 41) + (c * 1)
            pollution(loc - 39) = pollution(loc - 39) + (c * 1)
        End If
        If (loc + 1) Mod 40 < 39 And loc + 2 <= 1599 Then
            pollution(loc + 2) = pollution(loc + 2) + (c * 1)
       End If
        If (loc + 40) <= 1599 Then
            pollution(loc + 40) = pollution(loc + 40) + (c * 2)
        End If
        If (loc - 40) >= 0 Then
            pollution(loc - 40) = pollution(loc - 40) + (c * 2)
        End If
        If (loc + 80) <= 1599 Then
            pollution(loc + 80) = pollution(loc + 80) + (c * 1)
        End If
        If (loc - 80) >= 0 Then
            pollution(loc - 80) = pollution(loc - 80) + (c * 1)
        End If

    End If
    
End Function


Function makenoise(loc As Integer, c As Integer)
'c decides whether crime is made or cleaned
    If powergrid(loc) = 1 Then
        noise(loc) = noise(loc) + (c * 3)
        If (loc - 1) Mod 40 > 0 And loc - 2 >= 0 Then
            noise(loc - 2) = noise(loc - 2) + (c * 1)
        End If
        If (loc) Mod 40 > 0 And loc - 1 >= 0 And loc - 41 >= 0 And loc + 39 <= 1599 Then
            noise(loc - 1) = noise(loc - 1) + (c * 2)
            noise(loc - 41) = noise(loc - 41) + (c * 1)
            noise(loc + 39) = noise(loc + 39) + (c * 1)
        End If
        If (loc) Mod 40 < 39 And loc + 1 <= 1599 And loc + 41 <= 1599 And loc - 39 >= 0 Then
            noise(loc + 1) = noise(loc + 1) + (c * 2)
            noise(loc + 41) = noise(loc + 41) + (c * 1)
            noise(loc - 39) = noise(loc - 39) + (c * 1)
        End If
        If (loc + 1) Mod 40 < 39 And loc + 2 <= 1599 Then
            noise(loc + 2) = noise(loc + 2) + (c * 1)
       End If
        If (loc + 40) <= 1599 Then
            noise(loc + 40) = noise(loc + 40) + (c * 2)
        End If
        If (loc - 40) >= 0 Then
            noise(loc - 40) = noise(loc - 40) + (c * 2)
        End If
        If (loc + 80) <= 1599 Then
            noise(loc + 80) = noise(loc + 80) + (c * 1)
        End If
        If (loc - 80) >= 0 Then
            noise(loc - 80) = noise(loc - 80) + (c * 1)
        End If

    End If
    
End Function






Function makeFire(loc As Integer, c As Integer)
'c decides whether crime is made or cleaned
    If powergrid(loc) = 1 Then
        firezone(loc) = firezone(loc) + (c * 1)
        If (loc - 1) Mod 40 > 0 And loc - 2 >= 0 Then
            firezone(loc - 2) = firezone(loc - 2) + (c * 1)
        End If
        If (loc) Mod 40 > 0 And loc - 1 >= 0 And loc - 41 >= 0 And loc + 39 <= 1599 Then
            firezone(loc - 1) = firezone(loc - 1) + (c * 1)
            firezone(loc - 41) = firezone(loc - 41) + (c * 1)
            firezone(loc + 39) = firezone(loc + 39) + (c * 1)
        End If
        If (loc) Mod 40 < 39 And loc + 1 <= 1599 And loc + 41 <= 1599 And loc - 39 >= 0 Then
            firezone(loc + 1) = firezone(loc + 1) + (c * 1)
            firezone(loc + 41) = firezone(loc + 41) + (c * 1)
            firezone(loc - 39) = firezone(loc - 39) + (c * 1)
        End If
        If (loc + 1) Mod 40 < 39 And loc + 2 <= 1599 Then
            firezone(loc + 2) = firezone(loc + 2) + (c * 1)
       End If
        If (loc + 40) <= 1599 Then
            firezone(loc + 40) = firezone(loc + 40) + (c * 1)
        End If
        If (loc - 40) >= 0 Then
            firezone(loc - 40) = firezone(loc - 40) + (c * 1)
        End If
        If (loc + 80) <= 1599 Then
            firezone(loc + 80) = firezone(loc + 80) + (c * 1)
        End If
        If (loc - 80) >= 0 Then
            firezone(loc - 80) = firezone(loc - 80) + (c * 1)
        End If

    End If
    
End Function






Function propertyVal(loc As Integer)

propVal(RESID, loc) = 10
propVal(RESID, loc) = propVal(RESID, loc) - (5 * pollution(loc))
propVal(RESID, loc) = propVal(RESID, loc) - (2 * crime(loc))
propVal(RESID, loc) = propVal(RESID, loc) - (2 * firezone(loc))
propVal(RESID, loc) = propVal(RESID, loc) + (care(loc))
For a = 0 To 12
        If (loc + condvals(a)) >= 0 And loc + condvals(a) >= 0 Then
        If loc + condvals(a) <= 1599 Then
        If Abs(terrain(loc).Left - terrain(loc + condvals(a)).Left) < 751 Then
                If terrArr(loc + condvals(a)) = 1 Then
                    propVal(RESID, loc) = propVal(RESID, loc) + 4
                End If
                If terrArr(loc + condvals(a)) = -2 Then
                    propVal(RESID, loc) = propVal(RESID, loc) + 2
                End If
                If terrArr(loc + condvals(a)) = -1 Then
                    propVal(RESID, loc) = propVal(RESID, loc) - 1
                End If
        End If
        End If
        End If
Next a

propVal(COMME, loc) = 10
propVal(COMME, loc) = propVal(COMME, loc) - (pollution(loc))
For a = 0 To 12
        If (loc + condvals(a)) >= 0 And loc + condvals(a) >= 0 Then
        If loc + condvals(a) <= 1599 Then
        If Abs(terrain(loc).Left - terrain(loc + condvals(a)).Left) < 751 Then
                If terrArr(loc + condvals(a)) = 13 Then
                    propVal(COMME, loc) = propVal(COMME, loc) + 2
                End If
                If terrArr(loc + condvals(a)) = 2 Then
                    propVal(COMME, loc) = propVal(COMME, loc) + 1
                End If
        End If
        End If
        End If
Next a

propVal(INDUS, loc) = 10

propVal(INDUS, loc) = propVal(INDUS, loc) - (crime(loc))
For a = 0 To 12
        If (loc + condvals(a)) >= 0 And loc + condvals(a) >= 0 Then
        If loc + condvals(a) <= 1599 Then
        If Abs(terrain(loc).Left - terrain(loc + condvals(a)).Left) < 751 Then
                If terrArr(loc + condvals(a)) = 2 Then
                    propVal(INDUS, loc) = propVal(INDUS, loc) + 1
                End If
                If terrArr(loc + condvals(a)) = -1 Then
                    propVal(INDUS, loc) = propVal(INDUS, loc) + 2
                End If
                If terrArr(loc + condvals(a)) = 1 Then
                    propVal(INDUS, loc) = propVal(INDUS, loc) + 2
                End If
        End If
        End If
        End If
Next a


End Function


'   AND I THOUGHT I'D NEVER USE RECURSION!!!! OR DID I....

Function testpower(loc As Integer)
    If powergrid(loc) > 0 Then
        If loc - 40 >= 0 Then
            If terrArr(loc - 40) > 1 And powerchecked(loc - 40) = 0 Then
                If powergrid(loc - 40) <> 2 And terrArr(loc - 40) > 2 Then
                    powergrid(loc - 40) = 1
                End If
                powerchecked(loc - 40) = 1
                testpower (loc - 40)
            End If
        End If
        If loc + 40 < 1599 Then
            If terrArr(loc + 40) > 1 And powerchecked(loc + 40) = 0 Then
                If powergrid(loc + 40) <> 2 And terrArr(loc + 40) > 2 Then
                    powergrid(loc + 40) = 1
                End If
                powerchecked(loc + 40) = 1
                testpower (loc + 40)
            End If
        End If
        If loc Mod 40 > 0 And loc - 1 >= 0 Then
            If terrArr(loc - 1) > 1 And powerchecked(loc - 1) = 0 Then
                If powergrid(loc - 1) <> 2 And terrArr(loc - 1) > 2 Then
                    powergrid(loc - 1) = 1
                End If
                powerchecked(loc - 1) = 1
                testpower (loc - 1)
            End If
        End If
        If (loc + 1) Mod 40 > 0 And loc + 1 <= 1599 Then
            If terrArr(loc + 1) > 1 And powerchecked(loc + 1) = 0 Then
                If powergrid(loc + 1) <> 2 And terrArr(loc + 1) > 2 Then
                    powergrid(loc + 1) = 1
                End If
                powerchecked(loc + 1) = 1
                testpower (loc + 1)
            End If
        End If
    End If






End Function

Function havePower()
Dim a As Integer
Dim ptct As Integer
ptct = 0
    For a = 0 To 1599
        If powergrid(a) = 2 Then
            Call testpower(a)
            ptct = ptct + 1
            If ptct = numPlants Then
                a = 1600
            End If
        End If
    Next a

    For a = 0 To 1599
        If terrArr(a) > 1 And powergrid(a) = 0 Then
            If terrArr(a) = 3 Then
                terrain(a).Picture = noRes.Picture
            End If
            If terrArr(a) = 4 Then
                terrain(a).Picture = noCom.Picture
            End If
            If terrArr(a) = 5 Then
                terrain(a).Picture = noInd.Picture
            End If
            If terrArr(a) = 6 Then
                terrain(a).Picture = nopower.Picture
            End If
            If terrArr(a) = 8 Then
                terrain(a).Picture = noPolice.Picture
            End If
            If terrArr(a) = 9 Then
                terrain(a).Picture = noFire.Picture
            End If
            If terrArr(a) = 10 Then
                terrain(a).Picture = noHosp.Picture
            End If
            If terrArr(a) = 12 Then
                terrain(a).Picture = noPoweroad.Picture
            End If
            If terrArr(a) = 22 Then
                terrain(a).Picture = nopowerail.Picture
            End If
            If terrArr(a) = 16 Then
                terrain(a).Picture = noAir.Picture
            End If

        End If
        If terrArr(a) > 1 And powergrid(a) = 1 Then
            If terrArr(a) = 3 Then
                terrain(a).Picture = res.Picture
            End If
            If terrArr(a) = 4 Then
                terrain(a).Picture = com.Picture
            End If
            If terrArr(a) = 5 Then
                terrain(a).Picture = ind.Picture
            End If
            If terrArr(a) = 6 Then
                terrain(a).Picture = power.Picture
            End If
            If terrArr(a) = 8 Then
                terrain(a).Picture = police.Picture
            End If
            If terrArr(a) = 9 Then
                terrain(a).Picture = fire.Picture
            End If
            If terrArr(a) = 10 Then
                terrain(a).Picture = hospital.Picture
            End If
            If terrArr(a) = 12 Then
                terrain(a).Picture = poweroad.Picture
            End If
            If terrArr(a) = 22 Then
                terrain(a).Picture = powerail.Picture
            End If
            If terrArr(a) = 16 Then
                terrain(a).Picture = air.Picture
            End If
        End If
        powerchecked(a) = 0
    Next a

End Function



Function makecare(loc As Integer, c As Integer)
    If powergrid(loc) = 1 Then
        care(loc) = care(loc) + (c * 3)
        If (loc - 1) Mod 40 > 0 And loc - 1 >= 0 Then
            care(loc - 1) = care(loc - 1) + (c * 1)
        End If
        If (loc - 2) Mod 40 > 0 And loc - 2 >= 0 Then
            care(loc - 2) = care(loc - 2) + (c * 1)
        End If
        If (loc + 1) Mod 40 > 0 And loc + 1 <= 1599 Then
            care(loc + 1) = care(loc + 1) + (c * 1)
        End If
        If (loc + 2) Mod 40 > 0 And loc + 2 <= 1599 Then
            care(loc + 2) = care(loc + 2) + (c * 1)
       End If
        If (loc + 40) <= 1599 Then
            care(loc + 40) = care(loc + 40) + (c * 1)
        End If
        If (loc - 40) >= 0 Then
            care(loc - 40) = care(loc - 40) + (c * 1)
        End If
        If (loc + 80) <= 1599 Then
            care(loc + 80) = care(loc + 80) + (c * 1)
        End If
        If (loc - 80) >= 0 Then
            care(loc - 80) = care(loc - 80) + (c * 1)
        End If
            If ((loc - 41) + 1) Mod 40 > 0 And loc - 41 >= 0 Then
            care(loc - 41) = care(loc - 41) + (c * 1)
        End If
        If ((loc - 39) - 1) Mod 40 > 0 And loc - 39 >= 0 Then
            care(loc - 39) = care(loc - 39) + (c * 1)
        End If
        If ((loc + 39) + 1) Mod 40 >= 0 And loc + 39 <= 1599 Then
            care(loc + 39) = care(loc + 39) + (c * 1)
        End If
        If ((loc + 41) - 1) Mod 40 > 0 And loc + 41 <= 1599 Then
            care(loc + 41) = care(loc + 41) + (c * 1)
        End If
    End If
    
End Function

Function makeCrime(loc As Integer, c As Integer)
    If powergrid(loc) = 1 Then
        crime(loc) = crime(loc) + (c * 3)
        If (loc - 1) Mod 40 > 0 And loc - 1 >= 0 Then
            crime(loc - 1) = crime(loc - 1) + (c * 1)
        End If
        If (loc - 2) Mod 40 > 0 And loc - 2 >= 0 Then
            crime(loc - 2) = crime(loc - 2) + (c * 1)
        End If
        If (loc + 1) Mod 40 > 0 And loc + 1 <= 1599 Then
            crime(loc + 1) = crime(loc + 1) + (c * 1)
        End If
        If (loc + 2) Mod 40 > 0 And loc + 2 <= 1599 Then
            crime(loc + 2) = crime(loc + 2) + (c * 1)
       End If
        If (loc + 40) <= 1599 Then
            crime(loc + 40) = crime(loc + 40) + (c * 1)
        End If
        If (loc - 40) >= 0 Then
            crime(loc - 40) = crime(loc - 40) + (c * 1)
        End If
        If (loc + 80) <= 1599 Then
            crime(loc + 80) = crime(loc + 80) + (c * 1)
        End If
        If (loc - 80) >= 0 Then
            crime(loc - 80) = crime(loc - 80) + (c * 1)
        End If
            If ((loc - 41) + 1) Mod 40 > 0 And loc - 41 >= 0 Then
            crime(loc - 41) = crime(loc - 41) + (c * 1)
        End If
        If ((loc - 39) - 1) Mod 40 > 0 And loc - 39 >= 0 Then
            crime(loc - 39) = crime(loc - 39) + (c * 1)
        End If
        If ((loc + 39) + 1) Mod 40 >= 0 And loc + 39 <= 1599 Then
            crime(loc + 39) = crime(loc + 39) + (c * 1)
        End If
        If ((loc + 41) - 1) Mod 40 > 0 And loc + 41 <= 1599 Then
            crime(loc + 41) = crime(loc + 41) + (c * 1)
        End If
    End If
    
End Function

Function calcApproval()
Dim numlots As Integer
Dim numbuilt As Integer
Dim numtotal As Integer

numlots = 0
numbuilt = 0
numtotal = 0
approval = 0
    For a = 0 To 1599
        If terrArr(a) = 3 Then
            numlots = numlots + 1
        End If
        If terrArr(a) = 13 Then
            numbuilt = numbuilt + 1
        End If
    Next a
numtotal = numbuilt + numlots
If numtotal <> 0 Then
    approval = (100 * totalRes) / numtotal
End If

numlots = 0
numbuilt = 0
numtotal = 0
    For a = 0 To 1599
        If terrArr(a) = 4 Then
            numlots = numlots + 1
        End If
        If terrArr(a) = 14 Then
            numbuilt = numbuilt + 1
        End If
    Next a
numtotal = numbuilt + numlots
If numtotal <> 0 Then
    approval = (approval + ((100 * totalCom) / numtotal))
End If

numlots = 0
numbuilt = 0
numtotal = 0
    For a = 0 To 1599
        If terrArr(a) = 5 Then
            numlots = numlots + 1
        End If
        If terrArr(a) = 15 Then
            numbuilt = numbuilt + 1
        End If
        If terrArr(a) = 7 Then
            numbuilt = numbuilt + 1
        End If
    Next a
numtotal = numbuilt + numlots
If numtotal <> 0 Then
    approval = (approval + ((100 * totalInd) / numtotal)) / 3
End If

approval = approval + 5 - taxRate

Dim tmptot As Double
Dim tmpr As Double
Dim tmpc As Double
Dim tmpi As Double
tmptot = totalRes + totalCom + totalInd
If tmptot > 0 Then
    tmpr = 500 * (totalRes / tmptot)
    tmpc = 500 * (totalCom / tmptot)
    tmpi = 500 * (totalInd / tmptot)
End If
resbar.Width = 500 - tmpr
combar.Width = 500 - tmpc
indbar.Width = 500 - tmpi


End Function

Private Sub zoomIn_Click()
    HScroll1.Enabled = True
    VScroll1.Enabled = True
    HScroll1.Value = 0
    VScroll1.Value = 0
    Call invisible
    For a = 0 To 39
        For b = 0 To 39
            tmp = (a * 40) + b
            terrain(tmp).BorderStyle = 1
            terrain(tmp).Height = terrain(tmp).Height * 2.5
            terrain(tmp).Width = terrain(tmp).Width * 2.5
            terrain(tmp).Left = b * 360
            terrain(tmp).Top = a * 360
        Next b
    Next a
    Call revisible
    zoomIn.Enabled = False
    zoomOut.Enabled = True
End Sub

Private Sub zoomOut_Click()
    
HScroll1.Enabled = False
VScroll1.Enabled = False
    Call invisible
    For a = 0 To 39
        For b = 0 To 39
            tmp = (a * 40) + b
            terrain(tmp).BorderStyle = 0
            terrain(tmp).Height = terrain(tmp).Height / 2.5
            terrain(tmp).Width = terrain(tmp).Width / 2.5
            terrain(tmp).Left = b * 144
            terrain(tmp).Top = a * 144
        Next b
    Next a
    Call revisible
    zoomOut.Enabled = False
    zoomIn.Enabled = True
End Sub
