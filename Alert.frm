VERSION 5.00
Begin VB.Form Alert 
   BackColor       =   &H00FFFFFF&
   Caption         =   "eBay Bid Alert !"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8775
   Icon            =   "Alert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Alert.frx":0312
   ScaleHeight     =   2430
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   12000
      Left            =   240
      Top             =   1680
   End
   Begin VB.Label SellingLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$ Selling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   420
      Left            =   4560
      TabIndex        =   0
      Top             =   1200
      Width           =   3300
   End
End
Attribute VB_Name = "Alert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
 Result = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
 Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
 Unload Me
End Sub

