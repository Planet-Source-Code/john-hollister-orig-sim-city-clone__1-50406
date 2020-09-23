Attribute VB_Name = "Module1"
Global terrArr(0 To 1599) As Integer
'-3 = inspect
'-2 = forest
'-1 = rail
'0 = land
'1 = water
'2 = road
'3 = residential
'4 = commercial
'5 = industrial
'6 = powerline
'7 = powerplant
'8 = police
'9 = fire
'10 = hospital
'11
'12 = power&road
'13 = housed
'14 = built comm
'15 = built ind
'16 = airport
'17 = stadium
'26 = built airport

Global mapBuf(0 To 1599) As Integer



Global powergrid(0 To 1599) As Integer
Global powerchecked(0 To 1599) As Integer

Global funds As Long

Global monArray(1 To 12) As String

Global month As Integer
Global year As Integer

Global crime(0 To 1599) As Variant
Global firezone(0 To 1599) As Integer
Global care(0 To 1599) As Integer
Global pollution(0 To 1599) As Integer
Global noise(0 To 1599) As Integer

Global condvals(0 To 12) As Integer

Global population As Integer
Global jobs As Integer

Global approval As Integer
Global propVal(3 To 6, 0 To 1599) As Double
Global Const RESID = 3
Global Const COMME = 4
Global Const INDUS = 5
Global Const AIRPO = 6


Global taxRate As Double


Global sendRep As Integer


Global resRatio As Long
Global comRatio As Long
Global indRatio As Long

Global totalRes As Integer
Global totalCom As Integer
Global totalInd As Integer
Global totalAir As Integer


Global occupant(0 To 1599) As Integer
Global housetype(0 To 1599) As Integer

Global comJob(0 To 1599) As Integer
Global comtype(0 To 1599) As Integer

Global indtype(0 To 1599) As Integer
Global indjob(0 To 1599) As Integer

Global propstats(0 To 25) As Variant
