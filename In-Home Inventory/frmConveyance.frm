VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmConveyance 
   Caption         =   "Conveyance Bill"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   7815
   Begin MSDataGridLib.DataGrid dg 
      Height          =   1575
      Left            =   120
      TabIndex        =   33
      Top             =   3240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bill"
      Height          =   1575
      Left            =   5520
      TabIndex        =   34
      Top             =   3240
      Width           =   2175
      Begin VB.OptionButton optBill 
         Caption         =   "Total Paid Till Today"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton optBill 
         Caption         =   "Bill Of The Year"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optBill 
         Caption         =   "Bill Of The Month"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton optBill 
         Caption         =   "Bill Of The Day"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear"
      Height          =   375
      Left            =   6000
      TabIndex        =   32
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3960
      TabIndex        =   31
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6000
      TabIndex        =   30
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<<"
      Height          =   375
      Left            =   3960
      TabIndex        =   29
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   375
      Left            =   4920
      TabIndex        =   28
      Top             =   2400
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtClock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   360
      TabIndex        =   27
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   24
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   20
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Now"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   3495
      Begin VB.ComboBox cmbMedia 
         Height          =   315
         ItemData        =   "frmConveyance.frx":0000
         Left            =   1080
         List            =   "frmConveyance.frx":0016
         Sorted          =   -1  'True
         TabIndex        =   41
         Text            =   "Rickshaw"
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   40
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   960
         Width           =   2175
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   285
         Left            =   2400
         TabIndex        =   18
         Top             =   240
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   285
         Left            =   1080
         TabIndex        =   39
         Top             =   240
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tk."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   26
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   2040
         TabIndex        =   19
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Remarks:"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Amount:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Media:"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Destination:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "From"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Time:"
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.ComboBox cmbEmpName 
         Height          =   315
         ItemData        =   "frmConveyance.frx":0045
         Left            =   1080
         List            =   "frmConveyance.frx":0047
         Sorted          =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox cmbEmpId 
         Height          =   315
         ItemData        =   "frmConveyance.frx":0049
         Left            =   1080
         List            =   "frmConveyance.frx":004B
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   2775
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   1320
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
         Height          =   195
         Left            =   675
         TabIndex        =   6
         Top             =   1320
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Designation:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Employee ID:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL="
      Height          =   195
      Left            =   2520
      TabIndex        =   21
      Top             =   4800
      Width           =   615
   End
End
Attribute VB_Name = "frmConveyance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCon, rsEmp, rsRef, rsBill As New ADODB.Recordset
Dim sqlCon, sqlEmp As String
Public sqlBill As String
Dim Ed As Boolean

Public t1, t2, t3, t4, t5, t6, t7, t8, t9 As String

Private Sub chkBill_Click(Index As Integer)
'if chkbill(0).Caption =""
End Sub

Private Sub cmbEmpId_Change()
On Error GoTo ErrChange
If Trim(cmbEmpId.Text) = "" Then
cmdClearAll_Click
dg.ClearFields
cmbEmpName.Clear
Text1.Text = ""
MaskEdBox1.Text = "__/__/____"
Else
sqlEmp = "Select EmpID,EmpName,Designation FROM Employee WHERE EmpID=" & Trim(cmbEmpId.Text) ' & ""
End If
    
    Set rsEmp = cnnDB.Execute(sqlEmp)
    
    Dim i As Integer
If Not rsEmp.EOF Then
''''''''

''''''
   
    cmbEmpName.Text = rsEmp.Fields(1)
    Text1.Text = Trim(rsEmp.Fields(2))
    MaskEdBox1.Text = Date 'Format(Date, "dd / mm / yyyy")
    MaskEdBox2.Text = Format(Now, "Short time")
    MaskEdBox3.Text = Format(Now, "Short time")
'''''''''
'cmdPrev.Enabled = False
'cmdNext.Enabled = False
'cmdDel.Enabled = False
'cmdSave.Enabled = False
fncValidateFields
fncRefresh
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdDel.Enabled = True
cmdSave.Enabled = True
'sqlCon = "SELECT * FROM Conveyance c WHERE c.EmpID=" & t1 & "And c.Adate = #" & t9 & "# ORDER BY c.From_Time"
'Set rsCon = cnnDB.Execute(sqlCon)
sqlCon = "SELECT * FROM Conveyance c WHERE c.EmpID=" & t1 & "And c.Adate = #" & t9 & "# ORDER BY c.From_Time"
    'Set rsCon = New Recordset
  rsCon.Open sqlCon, cnnDB, adOpenStatic, adLockOptimistic
With rsCon
If Not .BOF Then
    .MoveFirst
End If
cmdPrev.Enabled = True
cmdNext.Enabled = True
cmdDel.Enabled = True
cmdSave.Enabled = True
    MaskEdBox2.Text = Format(.Fields(1).Value, "Short Time") 'from_time
    MaskEdBox3.Text = Format(.Fields(2).Value, "Short Time") ' to_time
    Text3.Text = .Fields(3).Value  'from
    Text4.Text = .Fields(4).Value ' dest
    cmbMedia.Text = .Fields(5).Value 'media
    Text6.Text = .Fields(6).Value 'amount
    Text7.Text = .Fields(7).Value 'remarks
    MaskEdBox1.Text = .Fields(8).Value 'ADate
End With
    fncRefresh
Else
cmbEmpName.Clear
Text1.Text = ""
MaskEdBox1.Text = "__/__/____"
End If
'cmdPrev.Enabled = True
'cmdNext.Enabled = True
'cmdDel.Enabled = True
'cmdSave.Enabled = True
Exit Sub
ErrChange:
  
End Sub

Private Sub cmbEmpId_Click()
cmbEmpId_Change
End Sub

Private Sub cmdAdd_Click()
On Error GoTo AddErr1

t1 = Trim(cmbEmpId.Text) 'EmpID
t2 = Trim(MaskEdBox2.Text) 'from_time
t3 = Trim(MaskEdBox3.Text) ' to_time
t4 = Trim(Text3.Text) 'from
t5 = Trim(Text4.Text) ' dest
t6 = Trim(cmbMedia.Text) 'media
t7 = Trim(Text6.Text) 'amount
t8 = Trim(Text7.Text) 'remarks
t9 = Trim(MaskEdBox1.Text) 'ADate
sqlCon = "INSERT INTO Conveyance (EmpId,From_Time,To_Time,From1,Dest,Media,Amount,Remarks,ADate)VALUES ( " & t1 & ",'" & t2 & "','" & t3 & "','" & t4 & "','" & t5 & "','" & t6 & "'," & t7 & ",'" & t8 & "','" & t9 & "' )"
'sqlCon = "INSERT INTO Conveyance (EmpId,From_Time,To_Time,From1)VALUES ( " & t1 & ",'" & t2 & "','" & t3 & "','" & t4 & "')" '& "'" & t5 & "','" & t6 & "'," & t7 & ",'" & t8 & "' )"
    Set rsCon = cnnDB.Execute(sqlCon)
    Text6.Text = 0# 'amount
    fncRefresh
    cmdClearAll_Click
Exit Sub
AddErr1:
  MsgBox " You Can't Add. Please check your Employee ID", , "Billing Error"
End Sub

Private Sub cmdClearAll_Click()
'cmdPrev.Enabled = False
'cmdNext.Enabled = False
'cmdDel.Enabled = False
'cmdSave.Enabled = False

    MaskEdBox2.Text = Format(Now, "Short Time") 'from_time
    MaskEdBox3.Text = Format(Now, "Short Time") ' to_time
    Text3.Text = ""  'from
    Text4.Text = "" ' dest
    cmbMedia.Text = "Rickshaw" 'media
    Text6.Text = Format(0, "##,##0.00")   'amount
    Text7.Text = "" 'remarks
    MaskEdBox1.Text = Format(Now, "Short Date") 'ADate
End Sub

Private Sub cmdDel_Click()
On Error GoTo DelErr1
'cmdPrev.Enabled = False
'cmdNext.Enabled = False
'cmdDel.Enabled = False
'cmdSave.Enabled = False
With rsCon
.Delete adAffectCurrent
End With
fncRefresh
Exit Sub
DelErr1:
    cmdDel.Enabled = False
    MsgBox " No Or Invalid Information", , "Delete"
End Sub

Private Sub cmdEdit_Click()
On Error GoTo EditErr1
'cmdPrev.Enabled = True
'cmdNext.Enabled = True
'cmdDel.Enabled = True
'cmdSave.Enabled = True
fncValidateFields
'    sqlCon = "SELECT * FROM Conveyance c WHERE c.EmpID=" & t1 & "And c.Adate = #" & t9 & "# ORDER BY c.From_Time"
'    'Set rsCon = New Recordset
'  rsCon.Open sqlCon, cnnDB, adOpenStatic, adLockOptimistic
With rsCon
If Not .BOF Or .EOF Then
    .MoveFirst
    cmbEmpId.Text = .Fields(0).Value   'EmpID
    MaskEdBox2.Text = Format(.Fields(1).Value, "Short Time") 'from_time
    MaskEdBox3.Text = Format(.Fields(2).Value, "Short Time") ' to_time
    Text3.Text = .Fields(3).Value  'from
    Text4.Text = .Fields(4).Value ' dest
    cmbMedia.Text = .Fields(5).Value 'media
    Text6.Text = .Fields(6).Value 'amount
    Text7.Text = .Fields(7).Value 'remarks
    MaskEdBox1.Text = .Fields(8).Value 'ADate
End If
End With
Exit Sub
EditErr1:
  MsgBox " Invalid Employee ID or Information Not Available", , "Billing Error"
'Command1.Enabled = True
End Sub

Private Sub cmdNext_Click()
On Error GoTo NextErr1
With rsCon
    .MoveNext
    If .EOF Then
        .MoveLast
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
    End If
    cmbEmpId.Text = .Fields(0).Value   'EmpID
    MaskEdBox2.Text = Format(.Fields(1).Value, "Short Time") 'from_time
    MaskEdBox3.Text = Format(.Fields(2).Value, "Short Time") ' to_time
    Text3.Text = .Fields(3).Value  'from
    Text4.Text = .Fields(4).Value ' dest
    cmbMedia.Text = .Fields(5).Value 'media
    Text6.Text = .Fields(6).Value 'amount
    Text7.Text = .Fields(7).Value 'remarks
    MaskEdBox1.Text = .Fields(8).Value 'ADate
End With
Exit Sub
NextErr1:
    'cmdDel.Enabled = False
    MsgBox " No Or Invalid Information", , "Next"
End Sub

Private Sub cmdPrev_Click()
On Error GoTo PrevErr1
With rsCon
    .MovePrevious
     If .BOF Then
        .MoveFirst
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
        Else
        
    cmbEmpId.Text = .Fields(0).Value   'EmpID
    MaskEdBox2.Text = Format(.Fields(1).Value, "Short Time") 'from_time
    MaskEdBox3.Text = Format(.Fields(2).Value, "Short Time") ' to_time
    Text3.Text = .Fields(3).Value  'from
    Text4.Text = .Fields(4).Value ' dest
    cmbMedia.Text = .Fields(5).Value 'media
    Text6.Text = .Fields(6).Value 'amount
    Text7.Text = .Fields(7).Value 'remarks
    MaskEdBox1.Text = .Fields(8).Value 'ADate
    End If
    
End With
Exit Sub
PrevErr1:
    'cmdDel.Enabled = False
    MsgBox " No Or Invalid Information", , "Previous"
End Sub

Private Sub cmdSave_Click()
'cmdPrev.Enabled = False
'cmdNext.Enabled = False
'cmdDel.Enabled = False
'cmdSave.Enabled = False
'If Ed = True Then

fncValidateFields
With rsCon
.Delete adAffectCurrent
.AddNew
    .Fields(0).Value = t1 'EmpID
    .Fields(1).Value = t2 'from_time
    .Fields(2).Value = t3 ' to_time
    .Fields(3).Value = t4 'from
    .Fields(4).Value = t5 ' dest
    .Fields(5).Value = t6 'media
    .Fields(6).Value = t7 'amount
    .Fields(7).Value = t8 'remarks
    .Fields(8).Value = t9 'ADate
 .UpdateBatch adAffectCurrent
' Ed = False
End With
'End If
fncRefresh
End Sub
Private Sub Command2_Click()
fncRefresh
End Sub
Private Sub Command3_Click()
End
End Sub
Private Sub Form_Load()
cmdPrev.Enabled = False
cmdNext.Enabled = False
cmdDel.Enabled = False
cmdSave.Enabled = False

Me.Height = 6000
Me.Width = 7935 '7320

fncGetEmpName
End Sub
Private Sub fncGetEmpName() 'type of user
sqlCon = "Select EmpID,EmpName FROM Employee"
    Set rsCon = cnnDB.Execute(sqlCon)
    cmbEmpName.Clear
    Dim i As Integer
    i = 0
            Do Until rsCon.EOF
            With cmbEmpId
                .AddItem (rsCon.Fields(0))
            End With
            With cmbEmpName
                .AddItem (rsCon.Fields(1))
            End With
        rsCon.MoveNext
        i = i + 1
        Loop
  cmbEmpName.Text = ""
End Sub
Public Sub fncRefresh()
On Error GoTo RefreshErr1
dg.Caption = "Refreshing........."
dg.ClearFields
fncValidateFields
cmdClearAll_Click
sqlCon = "SELECT c.From_Time,c.To_Time,c.Amount,c.From1,c.Dest,c.Media,c.Remarks FROM Conveyance c WHERE c.EmpID=" & t1 & "And c.Adate = #" & t9 & "# ORDER BY c.From_Time"
Set rsRef = cnnDB.Execute(sqlCon)
' If data found
    If Not (rsRef.BOF And rsRef.EOF) Then
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdDel.Enabled = True
        cmdSave.Enabled = True
        
        dg.Caption = "[ Date: " + t9 + " ],[ ID: " + t1 + " ], [ Name: " + cmbEmpName.Text + " ]"
        Set dg.DataSource = rsRef
        dg.Columns(0).Caption = "FROM(Time)"
        dg.Columns(1).Caption = "TO(Time)"
        dg.Columns(3).Caption = "FROM(Place)"
        dg.Columns(4).Caption = "DESTINATION(Place)"
        dg.Columns(5).Caption = "MEDIA(Vehicles)"
        dg.Columns(2).Alignment = dbgRight
        dg.Columns(2).Caption = "AMOUNT(Tk.)"
        dg.Columns(2).NumberFormat = "##,##0.00 Tk."
        dg.Columns(6).Caption = "Remarks"
        'Total Bill
        Dim i As Integer
        Dim Tk As Double
        i = 0
        Tk = 0#
        If Not rsRef.BOF Then
        rsRef.MoveFirst
        End If
        Do Until rsRef.EOF
            Tk = Tk + Val(rsRef.Fields(2))
            rsRef.MoveNext
            i = i + 1
        Loop
        Text2.Text = Tk 'get total bill
    'If no Data
    Else
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdDel.Enabled = False
        dg.Caption = "No Data Found"
        dg.ClearFields
        Set dg.DataSource = Nothing
        'dg.Columns(0).Text = "No Entry Available"
        Text2.Text = Format(0, "##,##0.00 Tk. Only")
    Exit Sub
End If
''''''''''''''
'Valid All Fields
    sqlCon = "SELECT * FROM Conveyance c WHERE c.EmpID=" & t1 & "And c.Adate = #" & t9 & "# ORDER BY c.From_Time"
    Set rsCon = New Recordset
    rsCon.Open sqlCon, cnnDB, adOpenStatic, adLockOptimistic
With rsCon
    If Not .BOF Or .EOF Then
        .MoveFirst

        cmbEmpId.Text = .Fields(0).Value   'EmpID
        MaskEdBox2.Text = Format(.Fields(1).Value, "Short Time") 'from_time
        MaskEdBox3.Text = Format(.Fields(2).Value, "Short Time") ' to_time
        Text3.Text = .Fields(3).Value  'from
        Text4.Text = .Fields(4).Value ' dest
        cmbMedia.Text = .Fields(5).Value 'media
        Text6.Text = .Fields(6).Value 'amount
        Text7.Text = .Fields(7).Value 'remarks
        MaskEdBox1.Text = .Fields(8).Value 'ADate
    End If
End With
''''''''''''''
Exit Sub
RefreshErr1:
'cmdPrev.Enabled = False
'cmdNext.Enabled = False
'cmdDel.Enabled = False
'cmdSave.Enabled = False
    'MsgBox " Invalid Employee ID or Information Not Available", , "Billing Error"
    cmbEmpId.SetFocus
End Sub

Private Sub optBill_Click(Index As Integer)
If optBill(0) Then
ElseIf optBill(1) Then
ElseIf optBill(2) Then
sqlBill = "SELECT c.From_Time,c.To_Time,c.Amount,c.From1,c.Dest,c.Media,c.Remarks FROM Conveyance c WHERE Year(c.ADate)=2003 ORDER BY c.EmpID"
'sqlBill = "SELECT sum(c.Amount)FROM Conveyance c WHERE Year(c.ADate)=2003 "
fncBill
ElseIf optBill(3) Then
End If

End Sub

Private Sub Text2_Change()
Text2.Text = Format(Val(Text2.Text), "##,##0.00 Tk. Only")
End Sub

Private Sub Text6_Change()
Text6.Text = Format(Val(Text6.Text), "##,##0.00")
End Sub

Private Sub Timer1_Timer()
Me.Caption = "[" + Format(Now, "Long Time") + "]"
  txtClock.Text = "TODAY: " + Format(Now, "Short Date")
End Sub
Public Sub fncValidateFields()
t1 = Trim(cmbEmpId.Text) 'EmpID
t2 = Trim(MaskEdBox2.Text) 'from_time
t3 = Trim(MaskEdBox3.Text) ' to_time
t4 = Trim(Text3.Text) 'from
t5 = Trim(Text4.Text) ' dest
t6 = Trim(cmbMedia.Text) 'media
t7 = Trim(Text6.Text) 'amount
t8 = Trim(Text7.Text) 'remarks
t9 = Trim(MaskEdBox1.Text) 'ADate
End Sub
Public Sub fncBill()
'On Error GoTo BillErr1
dg.Caption = "Refreshing........."
dg.ClearFields
fncValidateFields
cmdClearAll_Click
'sqlBill = "SELECT c.From_Time,c.To_Time,c.Amount,c.From1,c.Dest,c.Media,c.Remarks FROM Conveyance c WHERE c.EmpID=" & t1 & "And c.Adate = #" & t9 & "# ORDER BY c.From_Time"
Set rsBill = cnnDB.Execute(sqlBill)
' If data found
    If Not (rsBill.BOF And rsBill.EOF) Then
  
'dg.Caption = "[ Date: " + t9 + " ],[ ID: " + t1 + " ], [ Name: " + cmbEmpName.Text + " ]"
        Set dg.DataSource = rsBill
        dg.Columns(0).Caption = "FROM(Time)"
        dg.Columns(1).Caption = "TO(Time)"
        dg.Columns(3).Caption = "FROM(Place)"
        dg.Columns(4).Caption = "DESTINATION(Place)"
        dg.Columns(5).Caption = "MEDIA(Vehicles)"
        dg.Columns(2).Alignment = dbgRight
        dg.Columns(2).Caption = "AMOUNT(Tk.)"
        dg.Columns(2).NumberFormat = "##,##0.00 Tk."
        dg.Columns(6).Caption = "Remarks"
    'Total Bill
        Dim i As Integer
        Dim Tk As Double
        i = 0
        Tk = 0#
        If Not rsBill.BOF Then
        rsBill.MoveFirst
        End If
        Do Until rsBill.EOF
            Tk = Tk + Val(rsBill.Fields(2))
            rsBill.MoveNext
            i = i + 1
        Loop
        Text2.Text = Tk 'get total bill
    'If no Data
Else
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdDel.Enabled = False
        dg.Caption = "No Data Found"
        dg.ClearFields
        Set dg.DataSource = Nothing
        'dg.Columns(0).Text = "No Entry Available"
        Text2.Text = Format(0, "##,##0.00 Tk. Only")
    'Exit Sub
End If
''''''''''''''
'Valid All Fields
'    sqlCon = "SELECT * FROM Conveyance c WHERE c.EmpID=" & t1 & "And c.Adate = #" & t9 & "# ORDER BY c.From_Time"
'    Set rsCon = New Recordset
'    rsCon.Open sqlCon, cnnDB, adOpenStatic, adLockOptimistic
'With rsCon
'    If Not .BOF Or .EOF Then
'        .MoveFirst
'
'        cmbEmpId.Text = .Fields(0).Value   'EmpID
'        MaskEdBox2.Text = Format(.Fields(1).Value, "Short Time") 'from_time
'        MaskEdBox3.Text = Format(.Fields(2).Value, "Short Time") ' to_time
'        Text3.Text = .Fields(3).Value  'from
'        Text4.Text = .Fields(4).Value ' dest
'        cmbMedia.Text = .Fields(5).Value 'media
'        Text6.Text = .Fields(6).Value 'amount
'        Text7.Text = .Fields(7).Value 'remarks
'        MaskEdBox1.Text = .Fields(8).Value 'ADate
'    End If
'End With
''''''''''''''
Exit Sub
BillErr1:
'cmdPrev.Enabled = False
'cmdNext.Enabled = False
'cmdDel.Enabled = False
'cmdSave.Enabled = False
    MsgBox " Invalid Employee ID or Information Not Available", , "Billing Error"
    cmbEmpId.SetFocus
End Sub
