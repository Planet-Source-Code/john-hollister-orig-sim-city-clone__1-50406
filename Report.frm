VERSION 5.00
Begin VB.Form Report 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Property Report"
   ClientHeight    =   2325
   ClientLeft      =   5730
   ClientTop       =   4290
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label stat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label propvalue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label typeprop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label occup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Image display 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
display.Picture = form1.display.Picture

'display residential
If terrArr(sendRep) = 13 Or terrArr(sendRep) = 3 Then
    If terrArr(sendRep) = 13 Then
        If housetype(sendRep) < 3 Then
            typeprop.Caption = "Lower Class Shack"
        End If
        If housetype(sendRep) >= 3 Then
            typeprop.Caption = "Middle Class Castle"
        End If
    End If
    If terrArr(sendRep) = 3 Then
        typeprop.Caption = "Empty Residential Lot"
    End If
    occup.Caption = "Denizens: " & occupant(sendRep)
    If propVal(RESID, sendRep) >= 0 Then
        propvalue.Caption = "Property Value: $" & (propVal(RESID, sendRep) * 1000)
    Else
        propvalue.Caption = "Property Value: worthless"
    End If
    
    stat.Caption = "Fire: " & firezone(sendRep)

End If

'display commerce
If terrArr(sendRep) = 14 Or terrArr(sendRep) = 4 Then
    If terrArr(sendRep) = 14 Then
        If comtype(sendRep) < 3 Then
            typeprop.Caption = "Shit Hole Shop"
        End If
        If comtype(sendRep) >= 3 Then
            typeprop.Caption = "Conglomorate"
        End If
    End If
    If terrArr(sendRep) = 4 Then
        typeprop.Caption = "Empty Commercial Lot"
    End If
    occup.Caption = "Jobs: " & (comJob(sendRep) * 10)
    If propVal(COMME, sendRep) >= 0 Then
        propvalue.Caption = "Property Value: $" & (propVal(COMME, sendRep) * 1000)
    Else
        propvalue.Caption = "Property Value: worthless"
    End If
    stat.Caption = "Crime: " & crime(sendRep)

End If

'display industrial
If terrArr(sendRep) = 15 Or terrArr(sendRep) = 5 Then
    If terrArr(sendRep) = 15 Then
        If indtype(sendRep) < 3 Then
            typeprop.Caption = "Junk Warehouse"
        End If
        If indtype(sendRep) >= 3 Then
            typeprop.Caption = "Major Manufacturer"
        End If
    End If
    If terrArr(sendRep) = 5 Then
        typeprop.Caption = "Empty Industrial Lot"
    End If
    occup.Caption = "Jobs: " & (indjob(sendRep) * 10)
    If propVal(INDUS, sendRep) >= 0 Then
        propvalue.Caption = "Property Value: $" & (propVal(INDUS, sendRep) * 1000)
    Else
        propvalue.Caption = "Property Value: worthless"
    End If
    stat.Caption = "Pollution: " & pollution(sendRep)
End If







End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub

