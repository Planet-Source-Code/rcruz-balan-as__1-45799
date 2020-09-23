VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "balanca"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   Icon            =   "balanca.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim d As String
MSComm1.CommPort = 1
MSComm1.Settings = "9600,n,8,2"
MSComm1.InputLen = 10
MSComm1.PortOpen = True
MSComm1.Output = Chr$(13) & Chr$(10)
Do
DoEvents
buffer$ = buffer$ & MSComm1.Input
Loop Until Len(buffer$) >= 10
d = Mid$(buffer$, 2, 3) + "," + Mid$(buffer$, 6, 3)
Open "peso.txt" For Binary Access Write As #1
Put #1, 1, d
Close #1
MSComm1.PortOpen = False
End
End Sub
