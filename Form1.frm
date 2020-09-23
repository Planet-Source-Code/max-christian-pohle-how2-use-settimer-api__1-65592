VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SetTimer-API"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   1635
      Left            =   90
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   363
      TabIndex        =   4
      Top             =   2610
      Width           =   5505
   End
   Begin VB.CommandButton Command3 
      Caption         =   "add first timer again"
      Height          =   375
      Left            =   2700
      TabIndex        =   3
      Top             =   990
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kill first Timer"
      Height          =   375
      Left            =   2700
      TabIndex        =   2
      Top             =   540
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kill all Timers"
      Height          =   375
      Left            =   2700
      TabIndex        =   1
      Top             =   90
      Width           =   1635
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      Height          =   2355
      Left            =   90
      ScaleHeight     =   2295
      ScaleWidth      =   2385
      TabIndex        =   0
      Top             =   90
      Width           =   2445
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'written by Max Christian Pohle, http://www.coderonline.de/
'-----------------------------------------------------------
'   This demonstrates the use of the settimer-api
'   please don't care about this silly clock :-)
'-----------------------------------------------------------

Private Sub Command1_Click()
    KillAllTimers
End Sub

Private Sub Command2_Click()
    Kill_Timer Picture1, 1
End Sub

Private Sub Command3_Click()
    Set_Timer Picture1, 1, 500
End Sub

Private Sub Form_Load()
    Set_Timer Picture1, 1, 1
    Set_Timer Picture1, 999, 100
    Set_Timer Picture1, 122, 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillAllTimers
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Static Q As Double
    Const PI = 3.14159265
    
    'The Picturebox should be initalized in form_load
    'i wrote it here because its easier to understand how
    'the timers work when there is not so much code in form_load
    With Picture1
    .AutoRedraw = True
    .ScaleLeft = .ScaleWidth / -2
    .ScaleTop = .ScaleHeight / -2
    .ScaleHeight = 1
    .ScaleWidth = 1
    .DrawWidth = 4
    .FillStyle = 0
    .FillColor = vbBlue
    .FontTransparent = True
    
    
    'if not a key is pressed but this
    'event is called by the timer (value is negative)...
    If KeyCode < 0 Then
        Select Case -KeyCode
            Case 1      'first timer
                Q = (((PI * 400) / 60) * Format(Time, "ss")) / 200 - (PI / 2)
            
            Case 999    'second timer
                .Cls
                Picture1.Circle (0, 0), .ScaleWidth / 2.2, vbWhite
                Picture1.Line (0, 0)-(Cos((.Tag / 255) * PI * 2), Sin((.Tag / 255) * PI * 2)), RGB(255 - .Tag, .Tag, 0), BF
                Picture1.Line (0, 0)-(Cos(Q) / 2.5, Sin(Q) / 2.5), vbWhite
                .CurrentX = -.TextWidth(Format(Time, "ss")) / 2
                .CurrentY = -.TextHeight(Format(Time, "ss")) / 2
                Picture1.Print Format(Time, "ss")
            
            Case 122
                .DrawMode = vbMergePen
                .Tag = Val(.Tag) + 1
                If .Tag > 255 Then .Tag = 0
        
        End Select
    End If
    
    
    End With
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set_Timer Picture2, CLng(X), 500
    Picture2.Line (X, 0)-(X, Picture2.ScaleHeight), vbBlack
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Sgn(KeyCode) = -1 Then
        
        If Abs(KeyCode) <= Picture2.ScaleWidth Then
            X = Abs(KeyCode)
            Picture2.Line (X, 0)-(X, Picture2.ScaleHeight), vbWhite
            Kill_Timer Picture2, CLng(X)
        End If
        
    End If
End Sub

