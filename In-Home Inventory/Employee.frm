VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVBUTTONS.OCX"
Begin VB.Form frm_Employee 
   BackColor       =   &H00E0E0E0&
   Caption         =   "In-Home©"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   Icon            =   "Employee.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   8040
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   255
      Left            =   960
      TabIndex        =   30
      Top             =   4320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Add New"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "Employee.frx":2CCA
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
   Begin VB.Frame FrameFind 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Find By"
      ForeColor       =   &H00808080&
      Height          =   550
      Left            =   3480
      TabIndex        =   23
      Top             =   4080
      Width           =   4455
      Begin VB.OptionButton Option1ID 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Emp ID"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2Stat 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Emp St"
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option3Email 
         BackColor       =   &H00E0E0E0&
         Caption         =   "E-mail"
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option4Phone 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Phone"
         Height          =   255
         Left            =   3480
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
   End
   Begin LVbuttons.LaVolpeButton cmdDel 
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Delet"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "Employee.frx":2CE6
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
   Begin LVbuttons.LaVolpeButton cmdNext 
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   ">>>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "Employee.frx":2D02
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
   Begin VB.TextBox TextEid 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   14
      Top             =   3720
      Width           =   4455
   End
   Begin VB.TextBox TextEststus 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   11
      Top             =   3120
      Width           =   4455
   End
   Begin VB.TextBox TextEmail 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   5
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox TextCell 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5760
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox TextTnT 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox TextAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox TextFname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2830
      Left            =   240
      ScaleHeight     =   2775
      ScaleWidth      =   2295
      TabIndex        =   0
      Top             =   240
      Width           =   2350
      Begin VB.Image Image1 
         Height          =   2775
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2295
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin LVbuttons.LaVolpeButton cmdPrev 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "<<<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Employee.frx":2D1E
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
   Begin LVbuttons.LaVolpeButton cmdEdit 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "UP"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "Employee.frx":2D3A
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
   Begin LVbuttons.LaVolpeButton cmdLast 
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   3960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   ">>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
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
      MICON           =   "Employee.frx":2D56
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
   Begin LVbuttons.LaVolpeButton cmdFirst 
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "<<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
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
      MICON           =   "Employee.frx":2D72
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
   Begin LVbuttons.LaVolpeButton cmdAdd 
      Height          =   255
      Left            =   960
      TabIndex        =   20
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "Employee.frx":2D8E
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
   Begin LVbuttons.LaVolpeButton cmdfind 
      Height          =   255
      Left            =   960
      TabIndex        =   22
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Find"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
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
      MICON           =   "Employee.frx":2DAA
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
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Available Record:"
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
      Left            =   6240
      TabIndex        =   29
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Lblrecord 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C-No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7560
      TabIndex        =   28
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Cell"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone T && T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frm_Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

   Option Explicit
   
   Dim Cn As ADODB.Connection
   Dim rsrecord As New ADODB.Recordset
   Dim strConn As String
   Dim strSQL As String
   Dim FileLength As Long
   Dim Numblocks As Integer
   Dim LeftOver As Long
   Dim i As Integer
   Const BlockSize = 100000
                           

Private Sub cmdEdit_Click()
Dim ans As Variant

rsrecord("Ename") = TextFname.Text
rsrecord("Eaddress") = TextAddress.Text
rsrecord("Ecell") = TextCell.Text
rsrecord("Etnt") = TextTnT.Text
rsrecord("Eid") = TextEid.Text
rsrecord("Email") = TextEmail.Text
rsrecord("Estatus") = TextEststus.Text

rsrecord.Update

ans = MsgBox("Do You Want To Upgrade This Photo ID ?", vbYesNo, "Conferm To Upgrade")
        
        If ans = vbYes Then
             
            Call GetPic
            
            Else
            GoTo end_func
            
            End If

end_func:

End Sub



Private Sub cmdfind_Click()

If rsrecord.State = adStateOpen Then rsrecord.close

'************************find by Employee ID********************************

If Option1ID.Value = True Then

rsrecord.Open "Select ID,photo,Pname,description,Eaddress,Email,Ecell,Etnt,Ename,Estatus,Eid from Employee Where Eid like  '" & TextEid.Text & "'", Cn, adOpenKeyset, adLockOptimistic
Lblrecord.Caption = rsrecord.RecordCount

If rsrecord.RecordCount > 0 Then




TextFname.Text = rsrecord.Fields("Ename")
        TextAddress.Text = rsrecord.Fields("Eaddress")
        TextTnT.Text = rsrecord.Fields("Etnt")
        TextCell.Text = rsrecord.Fields("Ecell")
        TextEmail.Text = rsrecord.Fields("Email")
        TextEststus.Text = rsrecord.Fields("Estatus")
        TextEid.Text = rsrecord.Fields("Eid")

 FetchData
 Else
 MsgBox "No Available Information! ", vbInformation, "Not Available"
 End If

 Else
 
'************************Find by Employee Status********************************
 
If Option2Stat.Value = True Then


rsrecord.Open "Select ID,photo,Pname,description,Eaddress,Email,Ecell,Etnt,Ename,Estatus,Eid from Employee Where Estatus like  '" & TextEststus.Text & "'", Cn, adOpenKeyset, adLockOptimistic

Lblrecord.Caption = rsrecord.RecordCount

If rsrecord.RecordCount > 0 Then
   


TextFname.Text = rsrecord.Fields("Ename")
        TextAddress.Text = rsrecord.Fields("Eaddress")
        TextTnT.Text = rsrecord.Fields("Etnt")
        TextCell.Text = rsrecord.Fields("Ecell")
        TextEmail.Text = rsrecord.Fields("Email")
        TextEststus.Text = rsrecord.Fields("Estatus")
        TextEid.Text = rsrecord.Fields("Eid")
      
 FetchData
 Else
 MsgBox "No Available Information! ", vbInformation, "Not Available"
 End If
 
 Else
 
 
 
'************************Find by Employee Email********************************

 If Option3Email.Value = True Then


rsrecord.Open "Select ID,photo,Pname,description,Eaddress,Email,Ecell,Etnt,Ename,Estatus,Eid from Employee Where Email like  '" & TextEmail.Text & "'", Cn, adOpenKeyset, adLockOptimistic

Lblrecord.Caption = rsrecord.RecordCount

If rsrecord.RecordCount > 0 Then
   


TextFname.Text = rsrecord.Fields("Ename")
        TextAddress.Text = rsrecord.Fields("Eaddress")
        TextTnT.Text = rsrecord.Fields("Etnt")
        TextCell.Text = rsrecord.Fields("Ecell")
        TextEmail.Text = rsrecord.Fields("Email")
        TextEststus.Text = rsrecord.Fields("Estatus")
        TextEid.Text = rsrecord.Fields("Eid")
      
 FetchData
 Else
 MsgBox "No Available Information! ", vbInformation, "Not Available"
 End If
 
 Else
 
 
 
 
 
'************************Find by Employee Phone********************************

 If Option4Phone.Value = True Then


rsrecord.Open "Select ID,photo,Pname,description,Eaddress,Email,Ecell,Etnt,Ename,Estatus,Eid from Employee Where Etnt like  '" & TextTnT.Text & "'", Cn, adOpenKeyset, adLockOptimistic

Lblrecord.Caption = rsrecord.RecordCount

If rsrecord.RecordCount > 0 Then
   


TextFname.Text = rsrecord.Fields("Ename")
        TextAddress.Text = rsrecord.Fields("Eaddress")
        TextTnT.Text = rsrecord.Fields("Etnt")
        TextCell.Text = rsrecord.Fields("Ecell")
        TextEmail.Text = rsrecord.Fields("Email")
        TextEststus.Text = rsrecord.Fields("Estatus")
        TextEid.Text = rsrecord.Fields("Eid")
      
 FetchData
 Else
 MsgBox "No Available Information! ", vbInformation, "Not Available"
 End If
 
 Else
 
  
 MsgBox "Please Select Find Option ! ", vbInformation, "No Find Option"
 
 End If
 

 'MsgBox "Please Select Your Option ! ", vbInformation, "No Option"
 
 End If
 
 
 'MsgBox "Please Select Your Option ! ", vbInformation, "No Option"
 
 End If
 
 'MsgBox "Please Select Your Option ! ", vbInformation, "No Option"

 End If
 
 
 

End Sub

Private Sub cmdFirst_Click()
  On Error GoTo Error
    
    If rsrecord.RecordCount > 0 Then
        rsrecord.MoveFirst
        
        TextFname.Text = rsrecord.Fields("Ename")
        TextAddress.Text = rsrecord.Fields("Eaddress")
        TextTnT.Text = rsrecord.Fields("Etnt")
        TextCell.Text = rsrecord.Fields("Ecell")
        TextEmail.Text = rsrecord.Fields("Email")
        TextEststus.Text = rsrecord.Fields("Estatus")
        TextEid.Text = rsrecord.Fields("Eid")
        
        FetchData
    Else
        MsgBox "No records !!!!!!!!!"
    End If
    
Error:
    

End Sub

Private Sub CmdLast_Click()
  On Error GoTo Error
    If rsrecord.RecordCount > 0 Then
        rsrecord.MoveLast
        
        TextFname.Text = rsrecord.Fields("Ename")
        TextAddress.Text = rsrecord.Fields("Eaddress")
        TextTnT.Text = rsrecord.Fields("Etnt")
        TextCell.Text = rsrecord.Fields("Ecell")
        TextEmail.Text = rsrecord.Fields("Email")
        TextEststus.Text = rsrecord.Fields("Estatus")
        TextEid.Text = rsrecord.Fields("Eid")
        
        FetchData
        
    Else
        MsgBox "No records !!!!!!!!!"
    End If
    
Error:

End Sub

Private Sub cmdNext_Click()

  On Error GoTo Error
  
    If rsrecord.RecordCount > 0 Then
        rsrecord.MoveNext
        
        If rsrecord.EOF = True Then
            rsrecord.MoveLast
            
            cmdNext.Enabled = False
            cmdPrev.Enabled = True
            
            'MsgBox " Last Record"
            
            
        End If
    
        TextFname.Text = rsrecord.Fields("Ename")
        TextAddress.Text = rsrecord.Fields("Eaddress")
        TextTnT.Text = rsrecord.Fields("Etnt")
        TextCell.Text = rsrecord.Fields("Ecell")
        TextEmail.Text = rsrecord.Fields("Email")
        TextEststus.Text = rsrecord.Fields("Estatus")
        TextEid.Text = rsrecord.Fields("Eid")
        
        cmdPrev.Enabled = True
            
        FetchData
            
    Else
        
        MsgBox "No records"
    
    End If
    
Error:
    
End Sub

Private Sub CmdPrev_Click()

  On Error GoTo Error
  
    If rsrecord.RecordCount > 0 Then
        rsrecord.MovePrevious
        
        If rsrecord.BOF = True Then
            rsrecord.MoveFirst
            
            cmdNext.Enabled = True
            cmdPrev.Enabled = False
            
            
            'MsgBox "First record"
        End If
        
        TextFname.Text = rsrecord.Fields("Ename")
        TextAddress.Text = rsrecord.Fields("Eaddress")
        TextTnT.Text = rsrecord.Fields("Etnt")
        TextCell.Text = rsrecord.Fields("Ecell")
        TextEmail.Text = rsrecord.Fields("Email")
        TextEststus.Text = rsrecord.Fields("Estatus")
        TextEid.Text = rsrecord.Fields("Eid")
        
        'cmdPrev.Enabled = True
            
        cmdNext.Enabled = True
        
        FetchData
    
    Else
        
        MsgBox "No data !!!!!!!!"
    
    End If
Error:

End Sub

Private Sub CmdDel_Click()
    Dim ans As Variant
    
    If rsrecord.RecordCount > 0 Then
        
        
    ans = MsgBox("Are You Sure To Delet This Record ?", vbYesNo, "Conferm To Delet")
        
        If ans = vbYes Then
             
            rsrecord.Delete
            
            Else
            GoTo end_func
            
            End If
       
        
        If rsrecord.EOF = False Then
            cmdNext_Click
           
        Else
            CmdPrev_Click
           
        End If
    
    Else
        
        MsgBox "No records"
    End If

end_func:

End Sub

Private Sub Command1_Click()
Cn.close
Unload Me
End Sub


Private Sub Form_Load()
   
   Me.Top = frmMain.Top + frmMain.Height / 2 - 4500
   Me.Left = frmMain.Left + frmMain.Width / 2 - 6200

   If frm_user_pass.Text_User_ID.Text = "Administrator" Then
   
   frm_Employee.cmdAdd.Visible = True
   frm_Employee.cmdDel.Visible = True
   frm_Employee.cmdEdit.Visible = True
   frm_Employee.LaVolpeButton1.Visible = True

   Else
   
   frm_Employee.cmdAdd.Visible = False
   frm_Employee.cmdDel.Visible = False
   frm_Employee.cmdEdit.Visible = False
   frm_Employee.LaVolpeButton1.Visible = False
   End If
   
   

      Set Cn = New ADODB.Connection

     Cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Records.mdb;Persist Security Info=False;Jet OLEDB:Database Password=1234"


      If rsrecord.State = adStateOpen Then rsrecord.close

'Set rsrecord = cnnDB.Execute("Select ID,photo,Pname,description,Eaddress,Email,Ecell,Etnt,Ename,Estatus,Eid from Employee")
      
rsrecord.Open "Select ID,photo,Pname,description,Eaddress,Email,Ecell,Etnt,Ename,Estatus,Eid from Employee", Cn, adOpenKeyset, adLockOptimistic

Lblrecord.Caption = rsrecord.RecordCount


Me.Width = 8160
Me.Height = 5100


Call cmdFirst_Click

End Sub

Public Sub CmdAdd_Click()
Dim ans As Variant

rsrecord.AddNew

rsrecord("Ename") = TextFname.Text
rsrecord("Eaddress") = TextAddress.Text
rsrecord("Ecell") = TextCell.Text
rsrecord("Etnt") = TextTnT.Text
rsrecord("Eid") = TextEid.Text
rsrecord("Email") = TextEmail.Text
rsrecord("Estatus") = TextEststus.Text


'rsrecord.Update

ans = MsgBox("Do You Want To Add The Photo ?", vbYesNo, "Conferm To Add Photo")
        
If ans = vbYes Then
 LaVolpeButton1.Visible = True
 
Call GetPic

 Else
 
 LaVolpeButton1.Visible = True
 Call GetPic_NoPic
 
End If
    
    
   
End Sub

Private Sub FetchData()
      
    Dim J As Integer
       
        FileLength = CLng(rsrecord(2).Value)

        If FileLength > 10000 Then
            J = 5
        ElseIf FileLength < 10000 Then
            J = 10
        End If
      
'***************************************************
      
      Dim ByteData() As Byte
      Dim DestFileNum As Integer
      Dim DiskFile As String

      Me.MousePointer = vbHourglass
      
      DiskFile = App.Path & "\image1.bmp"
      If Len(Dir$(DiskFile)) > 0 Then
         Kill DiskFile
      End If

      DestFileNum = FreeFile
      Open DiskFile For Binary As DestFileNum
     
      Numblocks = FileLength / (BlockSize / J)
   

      
      For i = 1 To Numblocks
      
      ByteData() = rsrecord(1).GetChunk(BlockSize / J)
          Put DestFileNum, , ByteData()
      Next i
      Close DestFileNum

      Image1.Visible = True
      Image1.Picture = LoadPicture(App.Path & "\image1.bmp")
      
      Debug.Print "Complete"
      Me.MousePointer = vbNormal
   
End Sub

  


Private Sub GetPic()
   
      Dim PictBmp As String
      Dim ByteData() As Byte
      Dim SourceFile As Integer

      CommonDialog1.Filter = "(*.bmp;*.ico;*.jpg)|*.bmp;*.ico;*.jpg"
      CommonDialog1.ShowOpen
      
      If CommonDialog1.FileName <> "" Then
      PictBmp = CommonDialog1.FileName
      Me.MousePointer = vbHourglass
     
      rsrecord(3).Value = PictBmp
     
      SourceFile = FreeFile
      Open PictBmp For Binary Access Read As SourceFile

      FileLength = LOF(SourceFile)
      Debug.Print "Filelength is " & FileLength

      If FileLength = 0 Then

          Close SourceFile
          MsgBox PictBmp & " empty or not found."
          Exit Sub
      Else

          Numblocks = FileLength / BlockSize
          LeftOver = FileLength Mod BlockSize

          ReDim ByteData(LeftOver)
          Get SourceFile, , ByteData()
           
          rsrecord(1).AppendChunk ByteData()

          ReDim ByteData(BlockSize)
          For i = 1 To Numblocks
              Get SourceFile, , ByteData()
              rsrecord(1).AppendChunk ByteData()
          Next i

          rsrecord(2).Value = FileLength
 
          rsrecord.Update

         Close SourceFile
      End If

      Me.MousePointer = vbNormal
      FetchData
End If
End Sub

Private Sub LaVolpeButton1_Click()

TextFname.Text = ""
TextAddress.Text = ""
TextCell.Text = ""
TextTnT.Text = ""
TextEid.Text = ""
TextEmail.Text = ""
TextEststus.Text = ""


cmdNext.Enabled = False
cmdfind.Enabled = False
cmdFirst.Enabled = False
cmdPrev.Enabled = False
cmdLast.Enabled = False

LaVolpeButton1.Visible = False

End Sub



Private Sub TextEid_Change()
Option1ID.Value = True

End Sub

Private Sub TextEmail_Change()
Option3Email.Value = True
End Sub

Private Sub TextEststus_Change()
Option2Stat.Value = True

End Sub

Private Sub TextTnT_Change()
Option4Phone.Value = True

End Sub



Private Sub GetPic_NoPic()
   
      Dim PictBmp As String
      Dim ByteData() As Byte
      Dim SourceFile As Integer

'      CommonDialog1.Filter = "(*.bmp;*.ico;*.jpg)|*.bmp;*.ico;*.jpg"
'      CommonDialog1.ShowOpen

      'CommonDialog1.FileName = App.Path + "\In-Home Inventory\no.jpg" 'App.Path + "\no.jpg"
      'CommonDialog1.FileName = App.Path + "\In-Home Inventory\no.jpg" 'App.Path + "\no.jpg"
      
      CommonDialog1.FileName = "C:\Program Files\In-Home Inventory\no.jpg"  'App.Path + "\no.jpg"
      
      If CommonDialog1.FileName <> "" Then
      PictBmp = CommonDialog1.FileName
      Me.MousePointer = vbHourglass

      rsrecord(3).Value = PictBmp

      SourceFile = FreeFile
      Open PictBmp For Binary Access Read As SourceFile

      FileLength = LOF(SourceFile)
      Debug.Print "Filelength is " & FileLength

      If FileLength = 0 Then

          Close SourceFile
          MsgBox PictBmp & " empty or not found."
          Exit Sub
      Else

          Numblocks = FileLength / BlockSize
          LeftOver = FileLength Mod BlockSize

          ReDim ByteData(LeftOver)
          Get SourceFile, , ByteData()

          rsrecord(1).AppendChunk ByteData()

          ReDim ByteData(BlockSize)
          For i = 1 To Numblocks
              Get SourceFile, , ByteData()
              rsrecord(1).AppendChunk ByteData()
          Next i

          rsrecord(2).Value = FileLength

          rsrecord.Update

         Close SourceFile
      End If

      Me.MousePointer = vbNormal
      FetchData
End If



End Sub

