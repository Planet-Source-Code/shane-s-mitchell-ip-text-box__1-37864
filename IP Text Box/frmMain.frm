VERSION 5.00
Object = "*\AprjIPTextBox.vbp"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Text Box Demonstration Program"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin prjIPTextBox.IP IP1 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4095
   End
   Begin VB.Line Line2 
      X1              =   4320
      X2              =   4320
      Y1              =   240
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   4440
      X2              =   4440
      Y1              =   120
      Y2              =   2160
   End
   Begin VB.Label Label2 
      Caption         =   "IP Text Box by Shane Mitchell"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' /=============================================================
' IP Text Box
'
' Author: Shane Mitchell
' Version: 1.0
' Email: founder@extadi.com
'
' The best IP text box in the world.  It doesnt allow numbers
' over 255 in each segment no-matter how hard you try.  Also
' has great navigation around the IP Text Box.
' =============================================================/

Private Sub Form_Load()
    IP1_Change
End Sub

Private Sub IP1_Change()
    Label3.Caption = "IP:  " & IP1.Text & vbCrLf & _
        "Long IP:  " & IP1.GetLongIP
End Sub
