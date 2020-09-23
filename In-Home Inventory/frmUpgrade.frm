VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUpgrade 
   Caption         =   "Add/Upgrade Information"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   11865
   Begin VB.Frame frmIndentor 
      Caption         =   "Indentor's Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   43
      Top             =   2040
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Remarks 
         Height          =   765
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox textIndentor_Add 
         Height          =   765
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox textIndentor_Phone 
         Height          =   285
         Left            =   3000
         TabIndex        =   45
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox textIndentor_Name 
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label14 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   3000
         TabIndex        =   51
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Indentor's Address"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Indentor's Phone"
         Height          =   255
         Left            =   3000
         TabIndex        =   49
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Indentor's Name"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Types Of Product"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   32
      Top             =   360
      Width           =   5055
      Begin VB.OptionButton OptionFind 
         Caption         =   "Option1"
         Height          =   195
         Left            =   650
         TabIndex        =   57
         ToolTipText     =   "Select To Delete "
         Top             =   1250
         Width           =   170
      End
      Begin VB.CommandButton cmdfind 
         Caption         =   "Find"
         Enabled         =   0   'False
         Height          =   300
         Left            =   585
         TabIndex        =   56
         ToolTipText     =   "Click To Find Indentor's Information "
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton OptionReform 
         Caption         =   "Option1"
         Height          =   195
         Left            =   2810
         TabIndex        =   53
         ToolTipText     =   "Select To Upgrade "
         Top             =   1240
         Width           =   170
      End
      Begin VB.OptionButton OptionAdd 
         Caption         =   "Option1"
         Height          =   195
         Left            =   3920
         TabIndex        =   52
         ToolTipText     =   "Select To Add New"
         Top             =   1250
         Width           =   170
      End
      Begin VB.OptionButton OptionDel 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1720
         TabIndex        =   33
         ToolTipText     =   "Select To Delete "
         Top             =   1255
         Width           =   170
      End
      Begin VB.TextBox textSl_No 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Text            =   "Sl .No."
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox CombNameP 
         Height          =   315
         Left            =   1200
         TabIndex        =   38
         Text            =   "Product Name"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox textP_ID 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   37
         Text            =   "P .ID."
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdsave1 
         Caption         =   "Add "
         Enabled         =   0   'False
         Height          =   300
         Left            =   3840
         TabIndex        =   36
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton UP 
         Caption         =   "Reform"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2760
         TabIndex        =   35
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton CMDdEL2 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   34
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Serial No."
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Type of Pdoduct "
         Height          =   255
         Left            =   1200
         TabIndex        =   41
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Product ID"
         Height          =   255
         Left            =   3840
         TabIndex        =   40
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Customise Your System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton Command9 
         Caption         =   "Company's Stablishing"
         Height          =   330
         Left            =   2280
         TabIndex        =   25
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Company's  Web Site"
         Height          =   330
         Left            =   120
         TabIndex        =   24
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Company's  Telex"
         Height          =   330
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Company's    Fax"
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Company's Email"
         Height          =   330
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Your Phone && Etc"
         Height          =   330
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Your Company's  Address"
         Height          =   330
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Name of Title Bar"
         Height          =   330
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSDataGridLib.DataGrid dg1 
      Height          =   4575
      Left            =   5160
      TabIndex        =   12
      Top             =   2040
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8070
      _Version        =   393216
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   3600
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   6630
      Width           =   11655
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1296
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
      Caption         =   "Brand and Product"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5160
      TabIndex        =   0
      Top             =   360
      Width           =   6615
      Begin VB.OptionButton Option_Del 
         Caption         =   "Option1"
         Height          =   195
         Left            =   2450
         TabIndex        =   58
         ToolTipText     =   "Select To Delete "
         Top             =   1260
         Width           =   170
      End
      Begin VB.OptionButton OptionAdd1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   210
         TabIndex        =   55
         ToolTipText     =   "Select To Delete "
         Top             =   1260
         Width           =   170
      End
      Begin VB.OptionButton OptionReformed1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1370
         TabIndex        =   54
         ToolTipText     =   "Select To Delete "
         Top             =   1260
         Width           =   170
      End
      Begin VB.TextBox textParcent 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         MaxLength       =   6
         TabIndex        =   29
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton OptionDel2 
         Caption         =   "Option1"
         Height          =   90
         Left            =   0
         TabIndex        =   28
         ToolTipText     =   "Select To Delete "
         Top             =   6480
         Width           =   7710
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2400
         TabIndex        =   27
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpgrade 
         Caption         =   "Reform"
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         TabIndex        =   26
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "Close"
         Height          =   300
         Left            =   3490
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdsave2 
         Caption         =   "Add "
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox textquantity 
         Height          =   285
         Left            =   6000
         TabIndex        =   5
         Text            =   "Q.T."
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox textamount 
         Height          =   285
         Left            =   5160
         TabIndex        =   4
         Text            =   "BDT"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox textPID 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   3
         Text            =   "P.ID"
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cmdbrand 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox comProduct 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Percentage of Profit "
         Height          =   255
         Left            =   4800
         TabIndex        =   31
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   30
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Q.T"
         Height          =   255
         Left            =   6120
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "Cost BDT"
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "P .ID."
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Product Name"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Brand Name"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   0
      Picture         =   "frmUpgrade.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11940
   End
End
Attribute VB_Name = "frmUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, J, k As Integer

Dim t1, t2, t3, t4, t5, t6, t7, t8, t9, lisql As String

Dim flag As Boolean



Private Sub cmdclose_Click()

Me.Hide

End Sub

Private Sub cmdclose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdclose.Font.Bold = True
End Sub

Private Sub CmdDel_Click()
On Error GoTo err


lisql = "DELETE FROM Common Where Sl_No= " & textSl_No.Text & "and  P_ID='" & textPID.Text & "'"


Set rsrecord = cnnDB.Execute(lisql)


'***************EXECUTE QUARY TO GET DATA***************

Dim sqle As String

sqle = "Select Quantity from Common  Where Sl_No= " & textSl_No.Text ''& " order by P_ID"
Set rsrecord = cnnDB.Execute(sqle)
 
'cmdsave2.Enabled = False

Call Quantity_save
Text1.Text = 0


'**************CALL FUNCTION***************************

Call Product_From_Common
Call cmdRefresh
Call cmdRefreshDG


err:

End Sub

Private Sub cmdDel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdDel.Font.Bold = True

End Sub

Private Sub CMDdEL2_Click()

On Error GoTo err

lisql = "DELETE FROM Common_Product Where Sl_No= " & textSl_No.Text


Set rsrecord = cnnDB.Execute(lisql)

lisql = "DELETE FROM Common Where Sl_No= " & textSl_No.Text


Set rsrecord = cnnDB.Execute(lisql)


MsgBox "Deleate Success !"


'***************CALL FUNCTION*********************
Call Form_Load

'Call CombNameP_Click
'Call Product_From_Common
'Call cmdRefresh
'Call cmdRefreshDG

'**************DEL Option Button***************

OptionDel.Value = False
CMDdEL2.Enabled = False

err:


End Sub

Private Sub CMDdEL2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
CMDdEL2.Font.Bold = True
End Sub

Private Sub cmdfind_Click()
On Error GoTo err
Dim search As Variant

search = CombNameP.Text

If rsrecord.State = adStateOpen Then
rsrecord.close
End If


lisql = "select * from Common_Product where P_Name= '" & search & "'"

Set rsrecord = cnnDB.Execute(lisql)

textIndentor_Name.Text = rsrecord.Fields("Indentor_Name")
textIndentor_Add.Text = rsrecord.Fields("Indentor_Add")
textIndentor_Phone.Text = rsrecord.Fields("Indentor_Phone")
Remarks.Text = rsrecord.Fields("Indentor_Remarks")

Call CombNameP_Click

err:
End Sub

Private Sub cmdfind_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdfind.Font.Bold = True


End Sub

Private Sub cmdsave1_Click()
On Error GoTo err

Dim lisql As String

t1 = Trim(textP_ID.Text)
t2 = Trim(CombNameP.Text)
t3 = "0"
t4 = Trim(textIndentor_Name.Text + "")
t5 = Trim(textIndentor_Phone.Text + "")
t6 = Trim(textIndentor_Add.Text + "")
t7 = Trim(Remarks.Text + "")


lisql = "SELECT P_Name,P_ID FROM Common_Product Where P_Name= '" & CombNameP.Text & "' OR P_ID='" & textP_ID.Text & "'"

Set rsrecord = cnnDB.Execute(lisql)

If rsrecord.BOF Or rsrecord.EOF Then




lisql = "INSERT INTO Common_Product (P_ID,P_Name,Quantity,Indentor_Name,Indentor_Phone,Indentor_Add,Indentor_Remarks)VALUES ( '" & t1 & "','" & t2 & "','" & t3 & "','" & t4 & "','" & t5 & "','" & t6 & "','" & t7 & "')"

Set rsrecord = cnnDB.Execute(lisql)

cmdsave1.Enabled = False

''*************CALL COMBO FUNCTION******************

Call Form_Load


Else

MsgBox "Product Name Or P ID Alrady Exist", vbCritical, "Alrady Exist"

End If


err:


End Sub

Private Sub cmdsave1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdsave1.Font.Bold = True
End Sub

Private Sub cmdsave2_Click()
'On Error GoTo err

t1 = Trim(textSl_No.Text)
t2 = Trim(textPID.Text)
t3 = Trim(comProduct.Text)
t4 = Trim(cmdbrand.Text)
t5 = Trim(((Val(textParcent.Text) * Val(textamount.Text)) / 100) + Val(textamount.Text))
t6 = Trim(textquantity.Text)
t7 = Trim(textamount.Text)

Dim ans As String


If textParcent.Text = "" Then

ans = MsgBox("Do You Add Without Profit ?", vbYesNo, "Profit ! ")

If ans = vbYes Then
    
    GoTo CONTINUE
Else
    Exit Sub
End If

Else


CONTINUE:

If Val(textParcent.Text) <= 100 Then



'lisql = "SELECT Product,P_ID FROM Common Where Product= '" & comProduct.Text & "' OR P_ID='" & textPID.Text & "'"
lisql = "SELECT Product,P_ID FROM Common Where Product= '" & comProduct.Text & "' OR P_ID='" & textPID.Text & "' AND Sl_No= " & textSl_No.Text & " "

Set rsrecord = cnnDB.Execute(lisql)

If rsrecord.BOF Or rsrecord.EOF Then


lisql = "INSERT INTO Common (Sl_No,P_ID,Product,Brand,Amount_Tk,Quantity,Purchase_Tk)VALUES ( '" & t1 & "','" & t2 & "','" & t3 & "','" & t4 & "','" & t5 & "','" & t6 & "','" & t7 & "')"

Set rsrecord = cnnDB.Execute(lisql)


'******************************

Dim sqle As String

sqle = "Select Quantity from Common  Where Sl_No= " & textSl_No.Text ''& " order by P_ID"
Set rsrecord = cnnDB.Execute(sqle)
 
 
'cmdsave2.Enabled = False

Call Quantity_save

Text1.Text = 0


'**************CALL FUNCTION***************************

Call Product_From_Common
Call cmdRefresh
Call cmdRefreshDG

Else

MsgBox "Product Name Or P ID Alrady Exist", vbCritical, "Alrady Exist"

End If

Else

MsgBox "Parcentage Above 100", vbCritical, "Parcentage"

End If



End If
'err:
End Sub

Private Sub cmdsave2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdsave2.Font.Bold = True
End Sub

Private Sub cmdUpgrade_Click()
On Error GoTo err

Dim sql1 As String

t1 = Trim(textquantity.Text)
t2 = Trim(textamount.Text)
t3 = Trim(comProduct.Text)
t4 = Trim(cmdbrand.Text)

lisql = "SELECT Product FROM Common Where Product= '" & comProduct.Text & "' "

Set rsrecord = cnnDB.Execute(lisql)

If rsrecord.BOF Or rsrecord.EOF Then



sql1 = "update Common set Quantity='" & t1 & "',Amount_Tk='" & t2 & "',Product='" & t3 & "',Brand='" & t4 & "'Where Sl_No= " & textSl_No.Text & "and  P_ID='" & textPID.Text & "'"


cnnDB.Execute (sql1)

'cmdUpgrade.Enabled = False

'***************TO UPGRADE QUANTITY****************

Dim sqle As String

sqle = "Select Quantity from Common  Where Sl_No= " & textSl_No.Text ''& " order by P_ID"
Set rsrecord = cnnDB.Execute(sqle)
 
 
'***************CALL FUNCTION *********************
Call Quantity_save

Call Product_From_Common
Call cmdRefresh
Call cmdRefreshDG

Else

MsgBox "Product Name Alrady Exist Or Same Product", vbCritical, "Alrady Exist ! "

cmdUpgrade.Enabled = False
OptionReformed1.Value = False

End If

err:

End Sub

Private Sub cmdUpgrade_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdUpgrade.Font.Bold = True

End Sub

Private Sub CombNameP_Click()
On Error GoTo err

Dim search As Variant

search = CombNameP.Text

If rsrecord.State = adStateOpen Then
rsrecord.close
End If


lisql = "select * from Common_Product where P_Name= '" & search & "'"
Set rsrecord = cnnDB.Execute(lisql)
'Set dg.DataSource = rsrecord

rsrecord.MoveFirst

textSl_No.Text = rsrecord.Fields("SL_No")
textP_ID.Text = rsrecord.Fields("P_ID")
CombNameP.Text = rsrecord.Fields("P_Name")


'****************Gride Caption*******************

DG.Caption = CombNameP.Text
dg1.Caption = CombNameP.Text + " " + "List"


'***************CALL THE REFRESH FUNCTION***************

'Label4.Caption = "List of " & rsrecord.Fields("P_Name")

textParcent.Text = ""

cmdbrand.Clear

Call Product_From_Common
Call cmdRefreshDG
Call cmdRefresh

err:

End Sub

Private Sub Product_From_Common()
Dim i As Variant
Dim toolTips As String
Dim sql As String
Dim sql2 As String

Dim count As Integer

 sql = "select * from Common  where Sl_No= " & textSl_No.Text
 sql2 = sql + " order by P_ID"
 Set rsrecord = cnnDB.Execute(sql2)

 comProduct.Clear
 cmdbrand.Clear
 
 toolTips = count
 
 count = rsrecord.RecordCount
 For J = 1 To CInt(count)
 comProduct.AddItem rsrecord.Fields("Product") ' + "     On Hand " + rsrecord.Fields("Quantity")
 cmdbrand.AddItem rsrecord.Fields("Brand")   ' + "     On Hand " + rsrecord.Fields("Quantity")
  
 rsrecord.MoveNext
 
 Next
  
 comProduct.ToolTipText = "Total Item " & count
 
 rsrecord.MoveFirst
 comProduct.Text = rsrecord.Fields("Product") ' + "     On Hand " + rsrecord.Fields("Quantity")
 cmdbrand.Text = rsrecord.Fields("Brand")
 
End Sub



Private Sub comProduct_Click()
On Error GoTo err

Dim sql1 As String
Dim search As Variant

search = comProduct.Text

If rsrecord.State = adStateOpen Then
rsrecord.close
End If

sql1 = "select * from Common where Product= '" & search & " '"
Set rsrecord = cnnDB.Execute(sql1)

textPID.Text = rsrecord.Fields("P_ID")
textamount.Text = rsrecord.Fields("Amount_Tk")
textquantity.Text = rsrecord.Fields("Quantity")
cmdbrand.Text = rsrecord.Fields("Brand")

'************CALL FUNCTION*********************'

Call cmdRefreshDG
Call cmdRefresh

err:

End Sub



Private Sub DG_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

cmdDel.Font.Bold = False
UP.Font.Bold = False
CMDdEL2.Font.Bold = False
End Sub

Private Sub dg1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

cmdclose.Font.Bold = False


cmdDel.Font.Bold = False
UP.Font.Bold = False
CMDdEL2.Font.Bold = False

End Sub

Private Sub Form_Load()
On Error GoTo err

'Me.Top = fMainForm.Top + 100
'Me.Left = fMainForm.Left + 100

Me.Top = frmMain.Top + frmMain.Height / 2 - 4500
Me.Left = frmMain.Left + frmMain.Width / 2 - 6200


Me.Width = 12000
Me.Height = 7600

 

 Set rsrecord = cnnDB.Execute("Select * from Common_Product order by Sl_No asc")
  
 counter = rsrecord.RecordCount
 For i = 1 To CInt(counter)
 CombNameP.AddItem rsrecord.Fields("P_Name")
 rsrecord.MoveNext
 Next
   
 rsrecord.MoveFirst

 textP_ID.Text = rsrecord.Fields("P_ID")
 textSl_No.Text = rsrecord.Fields("Sl_No")
 CombNameP.Text = rsrecord.Fields("P_Name")

 '*************DG COL WIDTH ************************
 
 
 dg1.DefColWidth = 1250
 DG.DefColWidth = 1125

'***********************************************


flag = False


 
 '*************FUNCTION CALL************************
cmdbrand.Clear




Call CombNameP_Click
Call comProduct_Click
Call cmdRefreshDG

err:
End Sub



Private Sub cmdRefresh()
On Error GoTo err

Dim sqle As String

sqle = "Select C.Brand,C.Product,C.Amount_Tk,C.Quantity,C.P_ID,C.Purchase_Tk from Common C Where C.Sl_No= " & textSl_No.Text & " order by P_ID"
Set rsrecord = cnnDB.Execute(sqle)

Set dg1.DataSource = rsrecord

err:
End Sub

Private Sub cmdRefreshDG()

On Error GoTo err

Dim sqle As String

sqle = "Select P.Sl_No,P.P_ID,P.P_Name,P.Quantity from Common_Product P Where P.Sl_No= " & textSl_No.Text & " order by P_ID"
Set rsrecord = cnnDB.Execute(sqle)

Set DG.DataSource = rsrecord

err:
End Sub

Private Sub Quantity_save()

On Error GoTo err


For J = 1 To rsrecord.RecordCount

Text1.Text = Val(Text1.Text) + rsrecord.Fields("Quantity")

rsrecord.MoveNext

Next



t1 = Text1.Text

cnnDB.Execute " update Common_Product set Quantity='" & t1 & "' Where Sl_No=" & textSl_No.Text
 
Text1.Text = 0

err:

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdclose.Font.Bold = False

cmdsave2.Font.Bold = False
cmdsave1.Font.Bold = False
cmdUpgrade.Font.Bold = False
cmdfind.Font.Bold = False

cmdDel.Font.Bold = False
UP.Font.Bold = False
CMDdEL2.Font.Bold = False

End Sub


Private Sub Form_Resize()
'Me.Width = 12000
'Me.Height = 7600
'
End Sub



Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdclose.Font.Bold = False
cmdsave2.Font.Bold = False
cmdsave1.Font.Bold = False
cmdUpgrade.Font.Bold = False
cmdDel.Font.Bold = False
UP.Font.Bold = False
CMDdEL2.Font.Bold = False
cmdfind.Font.Bold = False
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdclose.Font.Bold = False

cmdsave2.Font.Bold = False
cmdsave1.Font.Bold = False
cmdUpgrade.Font.Bold = False
End Sub

Private Sub Command1_Click()

InputBox ("Edit Your Company Name ")

End Sub

Private Sub Command2_Click()
InputBox ("Edit Your Company Name ")
End Sub

Private Sub Command3_Click()

InputBox ("Edit Your Company's Address ")

End Sub

Private Sub Command4_Click()

InputBox ("Edit Your Company's Phone  ")

End Sub

Private Sub Command5_Click()

InputBox ("Edit Your Company's Email")

End Sub

Private Sub Command6_Click()

InputBox ("Edit Your Company's Fax")

End Sub

Private Sub Command7_Click()

InputBox ("Edit Your Company's Telex")

End Sub

Private Sub Command8_Click()

InputBox ("Edit Your Company's Web Site")

End Sub

Private Sub Command9_Click()

InputBox ("Edit Your Company's Stablishment")

End Sub







Private Sub frmIndentor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdfind.Font.Bold = False
End Sub

Private Sub Option_Del_Click()
 

 OptionReformed1.Value = False
 cmdUpgrade.Enabled = False

 cmdDel.Enabled = False
 OptionDel2.Value = False
    
 OptionAdd1.Value = False
 cmdsave2.Enabled = False
    
 Option_Del.Enabled = True
 cmdDel.Enabled = True


End Sub

Private Sub OptionAdd_Click()

On Error GoTo err

OptionDel.Value = False
CMDdEL2.Enabled = False
cmdfind.Enabled = False

OptionReform.Value = False
UP.Enabled = False

OptionAdd.Value = True
cmdsave1.Enabled = True
OptionFind.Value = False

 textSl_No.Enabled = False

textP_ID.Enabled = True
frmIndentor.Visible = True

textSl_No.Text = ""
CombNameP.Text = ""
textP_ID.Text = ""

cmdbrand.Text = ""
comProduct.Text = ""
textPID.Text = ""
textamount.Text = ""
textquantity.Text = ""

textPID.Enabled = True




rsrecord.MoveLast



err:
End Sub

Private Sub OptionAdd_DblClick()

OptionAdd.Value = False
cmdsave1.Enabled = False
frmIndentor.Visible = False

End Sub

Private Sub OptionAdd1_Click()
On Error GoTo err

    OptionReformed1.Value = False
    cmdUpgrade.Enabled = False

    cmdDel.Enabled = False
    OptionDel2.Value = False
    
    OptionAdd1.Value = True
    cmdsave2.Enabled = True
    
   
    textPID.Enabled = True
    
    cmdbrand.Text = ""
    comProduct.Text = ""
    textPID.Text = ""
    textamount.Text = ""
    textquantity.Text = ""

err:

End Sub

Private Sub OptionDel_Click()

   OptionDel.Value = True
   CMDdEL2.Enabled = True

   OptionReform.Value = False
   UP.Enabled = False
    
  OptionAdd.Value = False
  cmdsave1.Enabled = False
  frmIndentor.Visible = False

  OptionFind.Value = False
  cmdfind.Enabled = False
  
  textSl_No.Enabled = False

End Sub

Private Sub OptionDel_DblClick()
 
 OptionDel.Value = False
 CMDdEL2.Enabled = False

End Sub

Private Sub OptionDel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

OptionDel.ToolTipText = "Select To Delete" + " [" + CombNameP.Text + "] " + "Record"

End Sub









Private Sub OptionDel2_Click()

    OptionReformed1.Value = False
    cmdUpgrade.Enabled = False

    cmdDel.Enabled = True
    OptionDel2.Value = True
    
    OptionAdd1.Value = False
    cmdsave2.Enabled = False
    

End Sub

Private Sub OptionFind_Click()



OptionReform.Value = False
UP.Enabled = False

OptionAdd.Value = False
cmdsave1.Enabled = False

OptionDel.Value = False
CMDdEL2.Enabled = False


cmdfind.Enabled = True
OptionFind.Value = True

frmIndentor.Visible = True
textSl_No.Enabled = True

End Sub

Private Sub OptionReform_Click()

OptionReform.Value = True
UP.Enabled = True

cmdfind.Enabled = False
OptionAdd.Value = False
cmdsave1.Enabled = False

OptionDel.Value = False
CMDdEL2.Enabled = False

frmIndentor.Visible = False
OptionFind.Value = False

 textSl_No.Enabled = False
End Sub

Private Sub OptionReform_DblClick()
OptionReform.Value = False
UP.Enabled = False

End Sub

Private Sub OptionReformed1_Click()

OptionReformed1.Value = True
cmdUpgrade.Enabled = True

cmdDel.Enabled = False
OptionDel2.Value = False

OptionAdd1.Value = False
cmdsave2.Enabled = False
    

End Sub
Private Sub UP_Click()

Dim sql1 As String

On Error GoTo err
t1 = textP_ID.Text
t2 = CombNameP.Text


lisql = "SELECT P_Name FROM Common_Product Where P_Name= '" & CombNameP.Text & "' "

Set rsrecord = cnnDB.Execute(lisql)

If rsrecord.BOF Or rsrecord.EOF Then



sql1 = "update Common_Product set P_ID='" & t1 & "',P_Name='" & t2 & "'Where Sl_No= " & textSl_No.Text

cnnDB.Execute (sql1)


Call CombNameP_Click

Else

MsgBox "Product Name Or P ID Alrady Exist", vbCritical, "Alrady Exist"
UP.Enabled = False
OptionReform.Value = False
End If

err:
End Sub


Private Sub UP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

UP.Font.Bold = True

End Sub
