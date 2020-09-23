VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmReport 
   Caption         =   "All Report"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   6135
      Begin Crystal.CrystalReport CrAj 
         Left            =   1080
         Top             =   1680
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Submit"
         Height          =   255
         Left            =   4920
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   4680
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "To End Date"
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
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Start Date"
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
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "FrmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MaskEdBox1 = "__/__/____" Or MaskEdBox2 = "__/__/____" Then
        Exit Sub
    End If
differ ' It is date find function
CrAj.ReportFileName = App.Path + "\Emp_Atten_Date.rpt"
CrAj.Action = 0

End Sub
Private Sub differ()
On Error GoTo h1
CrAj.DataFiles(0) = App.Path + "\orient.mdb"
If MaskEdBox1 = "__/__/____" Then
  ElseIf Trim(MaskEdBox1) = "__/__/____" And Trim(MaskEdBox2) = "__/__/____" Then
        Exit Sub
    ElseIf Trim(MaskEdBox1) = "__/__/____" Then
            CrAj.SelectionFormula = "{EmpAttendence.ADate}=date(" & Year(MaskEdBox1) & "," & Month(MaskEdBox1) & "," & Day(MaskEdBox1) & ")"
        Else
            CrAj.SelectionFormula = "{EmpAttendence.ADate}= date(" & Year(MaskEdBox1) & "," & Month(MaskEdBox1) & "," & Day(MaskEdBox1) & ") to date(" & Year(MaskEdBox2) & "," & Month(MaskEdBox2) & "," & Day(MaskEdBox2) & ") "
    End If
     Exit Sub
h1:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "System"
    On Error GoTo 0
End Sub
Private Sub MaskEdBox2_GotFocus()
MaskEdBox2 = Date
End Sub
Private Sub MaskEdBox1_GotFocus()
MaskEdBox1 = Date
End Sub
