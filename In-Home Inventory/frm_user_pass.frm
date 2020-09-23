VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form frm_user_pass 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Login ..."
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   Icon            =   "frm_user_pass.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1150
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Login..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frm_user_pass.frx":0E42
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   26
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text_User_ID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      MaxLength       =   26
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unmask the Password"
      DragIcon        =   "frm_user_pass.frx":0E5E
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   1245
      Left            =   0
      Picture         =   "frm_user_pass.frx":70B8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frm_user_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

If Check1.Value = 1 Then
    Text1.PasswordChar = ""
Else
    Text1.PasswordChar = "*"
End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then
            On Error Resume Next
            Unload FRM_COMPANY
            Unload MDIForm1
            Unload frm_comp_name
            Unload frmSplash
    
                Dim x As Integer
                x = MsgBox("Are you sure you want to Close the Application ...", vbQuestion Or vbYesNo, "Want to Close Application ..")
                If x = 6 Then
                        Unload Me
                End If
    End If

End Sub

Private Sub Form_Load()

KeyPreview = True
 Me.Caption = "Login ..." + " " + Text_User_ID.Text

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
LaVolpeButton1.Font.Bold = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
            Unload frmMain
            Unload Me
End Sub
Private Sub LaVolpeButton1_Click()

Dim lisql As String
 
lisql = ("select * from `Identity` Where User_Id='" & LCase(Text_User_ID.Text) & "'and Passwd='" & Text1.Text & "'")

Set rsrecord = cnnDB.Execute(lisql)

If rsrecord.RecordCount > 0 Then

If LCase(rsrecord.Fields("User_Id")) = "administrator" Or LCase(rsrecord.Fields("User_Id")) = "admin" Then



frmMain.Profit_Report.Enabled = True
frmMain.New_User.Enabled = True
frmMain.Upgrade.Enabled = True
frmMain.User_Detail.Enabled = True





frmMain.Visible = True
frmMain.Show
Text1.Text = ""
Me.Hide

Else

frmMain.Profit_Report.Enabled = False
frmMain.New_User.Enabled = False
frmMain.Upgrade.Enabled = False
frmMain.User_Detail.Enabled = False



frmMain.Visible = True
frmMain.Show
Text1.Text = ""
Me.Hide

End If

Else

MsgBox "Password Or User ID Incorrect,Please Try Again", vbQuestion, "Password Incorrect "

End If

End Sub

Private Sub LaVolpeButton1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

LaVolpeButton1.Font.Bold = True

End Sub

Private Sub Text_User_ID_Change()

 Me.Caption = "Login ..." + " " + Text_User_ID.Text
End Sub

Private Sub Text_User_ID_Click()
Text1.Text = ""
End Sub

Private Sub Text_User_ID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call LaVolpeButton1_Click
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    Call LaVolpeButton1_Click
End If

End Sub

