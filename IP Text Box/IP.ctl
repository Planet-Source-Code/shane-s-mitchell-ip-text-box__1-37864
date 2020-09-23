VERSION 5.00
Begin VB.UserControl IP 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "0"
      Top             =   0
      Width           =   415
   End
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1035
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "0"
      Top             =   0
      Width           =   415
   End
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   525
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "0"
      Top             =   0
      Width           =   415
   End
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   0
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "0"
      Top             =   0
      Width           =   415
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   915
      TabIndex        =   3
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   405
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "IP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public Event Change()

Private Sub txt_Change(Index As Integer)
    RaiseEvent Change
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 3
        If txt(i).Text = "" Then txt(i).Text = "0"
    Next i
    
    txt(Index).SelStart = 0
    txt(Index).SelLength = Len(txt(Index).Text)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Shift = 2 And KeyCode = 86) Or (Shift = 1 And KeyCode = 45) Then
        KeyCode = 0
        Shift = 0
    End If
    
    If KeyCode = vbKeyLeft Then
        If txt(Index).SelStart = 0 Then
            If Index = 0 Then
                Beep
                KeyCode = 0
            Else
                txt(Index - 1).SetFocus
            End If
        End If
    ElseIf KeyCode = vbKeyRight Then
        If txt(Index).SelStart = Len(txt(Index).Text) Then
            If Index = 3 Then
                Beep
                KeyCode = 0
            Else
                txt(Index + 1).SetFocus
            End If
        End If
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 46 Then
            If Index < 3 Then
                txt(Index + 1).SetFocus
            Else
                Beep
            End If
            KeyAscii = 0
        ElseIf KeyAscii = 8 Then
            If txt(Index).Text = "" Or txt(Index).SelStart = 0 Then
                If Index = 0 Then
                    Beep
                Else
                    txt(Index - 1).SetFocus
                End If
                KeyAscii = 0
            End If
            If txt(Index).SelStart = 1 And Val(Mid(txt(Index).Text, 2, 1)) > 2 Then
                Beep
                KeyAscii = 0
            End If
            If txt(Index).SelStart = 2 And Val(Mid(txt(Index).Text, 3, 1)) > 5 And Mid(txt(Index).Text, 1, 1) = "2" Then
                Beep
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
        Exit Sub
    End If
    
    If KeyAscii > 50 And Len(txt(Index).Text) = txt(Index).SelLength Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
    If KeyAscii > 53 And txt(Index).SelStart = 1 And Mid(txt(Index).Text, 1, 1) = "2" Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
    If KeyAscii > 53 And txt(Index).SelStart = 2 And Mid(txt(Index).Text, 1, 2) = "25" Then
        KeyAscii = 0
        Beep
        Exit Sub
    End If
    
    If txt(Index).SelStart = 2 Then
        If Index < 3 Then
            txt(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = txt(3).Left + txt(3).Width
    UserControl.Height = txt(1).Height
End Sub

Public Property Get Text() As String
    Dim tst(0 To 3) As String
    tst(0) = txt(0).Text
    tst(1) = txt(1).Text
    tst(2) = txt(2).Text
    tst(3) = txt(3).Text
    
    If tst(0) = "" Then tst(0) = "0"
    If tst(1) = "" Then tst(1) = "0"
    If tst(2) = "" Then tst(2) = "0"
    If tst(3) = "" Then tst(3) = "0"
    
    Text = Join(tst, ".")
End Property

Public Property Let Text(tx As String)
    Dim sSpl() As String
    Dim i As Integer
    
    sSpl() = Split(tx, ".")
    If UBound(sSpl()) = 3 Then
        For i = 0 To 3
            If Not IsNumeric(sSpl(i)) Then
                Exit Property
            End If
        Next i
        For i = 0 To 3
            txt(i) = sSpl(i)
        Next i
    End If
End Property

Public Function GetLongIP() As Currency
    Dim tst(0 To 3) As String
    tst(0) = txt(0).Text
    tst(1) = txt(1).Text
    tst(2) = txt(2).Text
    tst(3) = txt(3).Text
    
    If tst(0) = "" Then tst(0) = "0"
    If tst(1) = "" Then tst(1) = "0"
    If tst(2) = "" Then tst(2) = "0"
    If tst(3) = "" Then tst(3) = "0"
    GetLongIP = CCur(tst(3)) + CCur(tst(2)) * 256 + _
        CCur(tst(1)) * 256 ^ 2 + CCur(tst(0)) * 256 ^ 3
End Function
