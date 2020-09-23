VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form EbayBidAdvisorForm 
   Caption         =   "eBay Bid Advisor "
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13485
   Icon            =   "eBayBidAdvisor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   13485
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox InfoScreen 
      Height          =   7935
      Index           =   1
      Left            =   3480
      ScaleHeight     =   7875
      ScaleWidth      =   13155
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   13215
      Begin VB.CommandButton Command7 
         Caption         =   "...BACK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9120
         TabIndex        =   21
         Top             =   6960
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   4695
         Left            =   240
         Picture         =   "eBayBidAdvisor.frx":08CA
         ScaleHeight     =   4635
         ScaleWidth      =   12555
         TabIndex        =   20
         Top             =   1800
         Width           =   12615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "NEXT..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11040
         TabIndex        =   15
         Top             =   6960
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Every 30 Seconds, this program will check if you have had any Bids on eBay.  It will Pop an Alert && Ring a Cash Register !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   13200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"eBayBidAdvisor.frx":A518C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   9
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   12720
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Once you have logged into eBay, please navigate to the My eBay screen as shown below:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   8
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   12585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"eBayBidAdvisor.frx":A523A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1200
         Index           =   6
         Left            =   360
         TabIndex        =   16
         Top             =   6600
         Width           =   8505
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox InfoScreen 
      Height          =   7935
      Index           =   0
      Left            =   7560
      ScaleHeight     =   7875
      ScaleWidth      =   13155
      TabIndex        =   5
      Top             =   -720
      Width           =   13215
      Begin VB.CommandButton Command5 
         Caption         =   "NEXT..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11040
         TabIndex        =   13
         Top             =   6960
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Download Source"
         Height          =   495
         Left            =   840
         TabIndex        =   12
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Absolutely! just press this button.  It is a hyperlink to a ZIP file and includes the complete Visual Basic 6.0 Source Code."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   720
         Index           =   4
         Left            =   840
         TabIndex        =   11
         Top             =   5160
         Width           =   12150
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Can I have the complete source code for this project for my peace of mind?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   360
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   4680
         Width           =   12060
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"eBayBidAdvisor.frx":A52E7
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   1440
         Index           =   2
         Left            =   840
         TabIndex        =   9
         Top             =   2880
         Width           =   12195
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Why every 30 Seconds?  Couldn't it be done faster?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   360
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   9540
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"eBayBidAdvisor.frx":A5427
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   12645
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Every 30 Seconds, this program will check if you have had any Bids on eBay.  It will Pop an Alert && Ring a Cash Register !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   13200
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "VIEW THE ALERT && RING CASH REGISTER"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   7560
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PAUSE TIMER"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1560
      Top             =   7440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START TIMER"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Timer StartTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   7440
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13215
      ExtentX         =   23310
      ExtentY         =   12938
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command8 
      Caption         =   "...BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   22
      Top             =   7560
      Width           =   1815
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   255
      Left            =   12480
      TabIndex        =   3
      Top             =   7680
      Width           =   495
      ExtentX         =   873
      ExtentY         =   450
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
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
      Left            =   7200
      TabIndex        =   23
      Top             =   7590
      Width           =   3300
   End
End
Attribute VB_Name = "EbayBidAdvisorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Command1.Enabled = False
 Command2.Enabled = True
 Timer1.Enabled = True
 Me.WindowState = 1
End Sub
Private Sub Command2_Click()
 Command1.Enabled = True
 Command2.Enabled = False
 Timer1.Enabled = False
End Sub
Private Sub Command3_Click()
 Alert.SellingLabel = SellingLabel
 Alert.Show
 WAVPlay "c:\Program Files\V8Software\CashReg.wav"
End Sub
Private Sub Command4_Click()
 WebBrowser2.Navigate "http://V8Software.com/eBayBidAdvisor.zip"
End Sub
Private Sub Command5_Click()
 InfoScreen(1).Visible = True
 InfoScreen(0).Visible = False
End Sub
Private Sub Command6_Click()
 InfoScreen(1).Visible = False
End Sub
Private Sub Command7_Click()
 InfoScreen(0).Visible = True
 InfoScreen(1).Visible = False
End Sub
Private Sub Command8_Click()
 InfoScreen(1).Visible = True
End Sub
Private Sub Form_Load()
 If App.PrevInstance Then End
 InfoScreen(0).Top = 120
 InfoScreen(0).Left = 120
 InfoScreen(1).Top = 120
 InfoScreen(1).Left = 120
 Close
 On Error Resume Next
 MkDir "c:\Program Files"
 MkDir "c:\Program Files\V8Software\"
 If Dir("c:\Program Files\V8Software\CashReg.wav") = "" Then
  Call BuildTheWavFile
 End If
 Open "c:\Program Files\V8Software\Init.HTM" For Output As #1
 Print #1, "Calling the eBay.com web page..."
 Close
 DoEvents
 WebBrowser1.Navigate "c:\Program Files\V8Software\Init.HTM"
 DoEvents
 frmSplash.Show vbModal, Me
 StartTimer.Enabled = True
End Sub
Private Sub SellingLabel_Change()
 Call Command3_Click
End Sub
Private Sub StartTimer_Timer()
 StartTimer.Enabled = False
 WebBrowser1.Navigate "http://my.ebay.com/"
End Sub
Private Sub Timer1_Timer()
 WebBrowser1.Refresh
End Sub
Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
 On Error Resume Next
 a$ = UCase$(WebBrowser1.Document.DocumentElement.innertext) & Space$(5000)
 If InStr(a$, "WILL SELL") > 0 And InStr(a$, "AMOUNT:") > 0 Then
  S = InStr(a$, "WILL SELL")
  If S > 0 Then
    D = InStr(S, a$, "AMOUNT:")
  End If
  If D > 0 Then
   b$ = Mid$(a$, D, 40)
   S = InStr(b$, "FIXED")
   If S Then
    b$ = Left$(b$, S - 1)
    b$ = Trim$(b$) & Space$(40)
   End If
   b$ = Trim$(b$)
   If LastSelling$ <> b$ Then
    SellingLabel.Caption = b$
   End If
   LastSelling$ = b$
  End If
 End If
End Sub
Sub WAVPlay(WavFile As String)
 SND_ASYNC = &H1
 SND_NODEFAULT = &H2
 wFlags% = SND_ASYNC Or SND_NODEFAULT
 Ignore = sndPlaySound(WavFile, wFlags%)
End Sub
