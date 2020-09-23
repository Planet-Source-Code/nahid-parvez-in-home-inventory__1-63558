VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form changepasswoard 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3720
   ClientLeft      =   3660
   ClientTop       =   1665
   ClientWidth     =   4950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   4950
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   4695
      Begin VB.CommandButton Command2 
         Caption         =   "E&xit"
         Height          =   300
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4695
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   3840
         Top             =   840
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox Text5 
         Height          =   195
         Left            =   3600
         TabIndex        =   13
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   615
         TabIndex        =   5
         Top             =   1440
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "New User Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   855
         TabIndex        =   4
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Current Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Current User Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "changepasswoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "You must Input All Fild", vbInformation, "Madani"
Text1.SetFocus
Else
Adodc1.RecordSource = "select * from Emp where EmpID='" & Trim(Text5.Text) & "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
Adodc1.Recordset.AddNew
End If
With Adodc1.Recordset
!EmpID = Trim(Text3.Text)
!Password = Trim(Text4.Text)
.Update
End With
MsgBox "Data has been saved", vbInformation, "System Madani"
Command1.Enabled = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End If
End Sub

Private Sub Command2_Click()
Unload Me
Set changepasswoard = Nothing
End Sub

Private Sub Form_Load()
databaseconnection
Adodc1.connectionstring = connectionstring
End Sub

Private Sub Text1_GotFocus()
Command1.Enabled = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Text2_GotFocus()
Command1.Enabled = False
End Sub

Private Sub Text2_LostFocus()
Dim user As String
Dim pass As String

Adodc1.RecordSource = "select * from Emp where EmpID='" & Trim(Text5.Text) & "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then

With Adodc1.Recordset
user = Trim(!EmpID)
pass = Trim(!Password)
End With
If Trim(Text1.Text) = Trim(user) And Trim(Text2.Text) = Trim(pass) Then
Text3.SetFocus
Command1.Enabled = True
Else
MsgBox "Invalid Password", vbInformation, "System Madani"
Text1.SetFocus
End If
End If
End Sub
