VERSION 5.00
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "ANIGIF.OCX"
Begin VB.Form About 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "About.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin AniGIFCtrl.AniGIF AniGIF1 
      Height          =   1695
      Left            =   480
      TabIndex        =   5
      Top             =   -360
      Width           =   3375
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "About.frx":2F84A
      ExtendWidth     =   5953
      ExtendHeight    =   2990
   End
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   240
      Top             =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel: 011-011539"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1860
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Web:http/www.in-homebd.netfirms.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "E@mail:nahidparvez@yahoo.com "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Khandoker Nahid  Parvez  "
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
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   2775
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'
'Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'
'Private Const GWL_EXSTYLE = (-20)
'Private Const WS_EX_LAYERED = &H80000
'Private Const LWA_ALPHA = &H2&

Private Sub Form_Load()

'Me.Top = fMainForm.Top + 3000
'Me.Left = fMainForm.Left + 4000
'

Me.Top = frmMain.Top + frmMain.Height / 2 - 2000 '    + 3000
Me.Left = frmMain.Left + frmMain.Width / 2 - 2000 '  4000


''Flash.Visible = False
'Flash.Movie = App.Path & "\Flash\123.swf"

'*******************TransParet Window************************************
''Dim LEVEL As Byte
''LEVEL = 180 '255 5
''Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
''
''Call SetLayeredWindowAttributes(Me.hwnd, 0, LEVEL, LWA_ALPHA)

End Sub

Private Sub Timer1_Timer()

Me.Hide

End Sub


