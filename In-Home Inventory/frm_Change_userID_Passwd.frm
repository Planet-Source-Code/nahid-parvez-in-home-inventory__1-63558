VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form frm_Change_userid_passwd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customise Form"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   Icon            =   "frm_Change_userID_Passwd.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User"
      DragIcon        =   "frm_Change_userID_Passwd.frx":0E42
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
      Left            =   4080
      TabIndex        =   18
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox CheckAdministator 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Administrator"
      DragIcon        =   "frm_Change_userID_Passwd.frx":709C
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
      TabIndex        =   17
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Operator"
      DragIcon        =   "frm_Change_userID_Passwd.frx":D2F6
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
      Left            =   2280
      TabIndex        =   16
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text_FName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      MaxLength       =   26
      TabIndex        =   14
      Top             =   120
      Width           =   2415
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   300
      Left            =   3720
      TabIndex        =   13
      Top             =   1365
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "Customise Button"
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
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frm_Change_userID_Passwd.frx":13550
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
   Begin VB.TextBox TextComants 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   26
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox TextMdate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   26
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox TextCdate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   26
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButtonPrev 
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "<<<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frm_Change_userID_Passwd.frx":1356C
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
   Begin LVbuttons.LaVolpeButton LaVolpeButtonNext 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   ">>>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frm_Change_userID_Passwd.frx":13588
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
   Begin VB.TextBox Text_User_ID 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      MaxLength       =   26
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unmask the Password"
      DragIcon        =   "frm_Change_userID_Passwd.frx":135A4
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
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   26
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "User Level Permission:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Comants"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Modify Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Create Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
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
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   1365
      Left            =   0
      Picture         =   "frm_Change_userID_Passwd.frx":197FE
      Stretch         =   -1  'True
      Top             =   0
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
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frm_Change_userid_passwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lisql, t1 As String

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
 

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
LaVolpeButton1.Font.Bold = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    On Error Resume Next
'            Unload frmMain
'            Unload Me
'
End Sub

Private Sub LaVolpeButton1_Click()


If Text1.Text = "" Or Text_User_ID.Text = "" Then

Me.Caption = "Faill To Add New User "

MsgBox "Password Or User ID Empty,Please Try Again", vbInformation, "Password Or User ID Empty ! "

Else

If LaVolpeButton1.Caption = "Add New User" Then

'lisql = "INSERT INTO  `Identity`  (User_Id,Passwd) VALUES ( '" & Text_User_ID.Text & "','" & Text1.Text & "')"

t1 = "OK"

lisql = "INSERT INTO  `Identity`  (User_Id,Passwd,CDate,MDate,Comants,FName) VALUES ( '" & Text_User_ID.Text & "','" & Text1.Text & "','" & Date & "','" & Date & "','" & t1 & "','" & Text_FName.Text & "')"

Set rsrecord = cnnDB.Execute(lisql)
Me.Caption = "Add New User Success "
Text1.Text = ""
Text_User_ID.Text = ""
Text_FName.Text = ""

End If


'*******************************************************************

If LaVolpeButton1.Caption = "Change" Then

lisql = "update `Identity` set User_ID='" & Text_User_ID.Text & "',MDate='" & Date & "' Where Passwd like '" & Text1.Text & "'"

cnnDB.Execute (lisql)

Me.Caption = "Reforme User ID Success "


frmMain.Hide
Me.Hide
frm_user_pass.Show



'Else
'
'MsgBox "Change User ID Impossible,Please Try Again", vbInformation, "ID Change Error! "

End If



'*****************************************************************

If LaVolpeButton1.Caption = "OK" Then


lisql = "update `Identity` set Passwd='" & Text1.Text & "'Where User_ID like '" & Text_User_ID.Text & "'"

cnnDB.Execute (lisql)

Me.Caption = "Reforme User Password Success "

'Else
'
'MsgBox "Change User Password Impossible,Please Try Again", vbInformation, "Password Change Error! "

End If

End If


End Sub

Private Sub LaVolpeButton1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

LaVolpeButton1.Font.Bold = True

End Sub

Private Sub LaVolpeButtonNext_Click()


If rsrecord.RecordCount > 0 Then

rsrecord.MoveNext

If rsrecord.EOF Then

LaVolpeButtonNext.Enabled = False
LaVolpeButtonPrev.Enabled = True

Me.Caption = "End Of The Record !"

Else

LaVolpeButtonPrev.Enabled = True

Text_User_ID.Text = rsrecord.Fields("User_ID")
Text1.Text = rsrecord.Fields("Passwd")
TextCdate.Text = rsrecord.Fields("Cdate")
TextMdate.Text = rsrecord.Fields("Mdate")
TextComants.Text = rsrecord.Fields("Comants")
Text_FName.Text = rsrecord.Fields("FName")


End If

Else
MsgBox "No Record"

End If



End Sub

Private Sub LaVolpeButtonPrev_Click()

If rsrecord.RecordCount > 0 Then

rsrecord.MovePrevious

If rsrecord.BOF Then

LaVolpeButtonNext.Enabled = True
LaVolpeButtonPrev.Enabled = False

Me.Caption = "End Of The Record !"

Else

LaVolpeButtonNext.Enabled = True

Text_User_ID.Text = rsrecord.Fields("User_ID")
Text1.Text = rsrecord.Fields("Passwd")
TextCdate.Text = rsrecord.Fields("Cdate")
TextMdate.Text = rsrecord.Fields("Mdate")
TextComants.Text = rsrecord.Fields("Comants")
Text_FName.Text = rsrecord.Fields("FName")


End If

Else
MsgBox "No Record"

End If



End Sub

Private Sub Text_User_ID_Change()
 Me.Caption = "Login ..." + Text_User_ID.Text
 
End Sub

Private Sub Text_User_ID_Click()
'Text1.Text = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call LaVolpeButton1_Click
End If
End Sub
