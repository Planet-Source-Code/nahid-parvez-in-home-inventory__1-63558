VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Client 
   Caption         =   "Client Information "
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   255
      Left            =   960
      TabIndex        =   33
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   255
      Left            =   2040
      TabIndex        =   32
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<<<<<<"
      Height          =   255
      Left            =   3600
      TabIndex        =   30
      Top             =   6150
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">>>>>>>"
      Height          =   255
      Left            =   5040
      TabIndex        =   29
      Top             =   6150
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   24
      Top             =   6150
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   23
      Top             =   6150
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DGProduct 
      Height          =   3855
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6800
      _Version        =   393216
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
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   10200
      TabIndex        =   21
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Find Now"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   6360
      Width           =   9735
      Begin VB.CommandButton Command4 
         Caption         =   "Execute"
         Height          =   375
         Left            =   2760
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Del On Basis of date"
         Height          =   375
         Left            =   4080
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdDelet 
         Caption         =   "Delet"
         Height          =   315
         Left            =   7080
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OptionDate 
         Caption         =   "Find Date"
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "q"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8400
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Optinv 
         Caption         =   "Invoice No"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Client Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9855
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   7800
         TabIndex        =   34
         Text            =   "Text3"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   6000
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4320
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox texttotal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox textInvoice_ID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox textCphone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox textCfax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox textCname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox textCmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox textCaddress 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox textDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5880
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Total Amount BDT:"
         Height          =   255
         Left            =   7800
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Fax :"
         Height          =   255
         Left            =   5880
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Email :"
         Height          =   255
         Left            =   7800
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Phone :"
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Date:"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   6
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Invoice No:"
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Image Image2 
      Height          =   6975
      Left            =   9960
      Picture         =   "frm_Client.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frm_Client.frx":92C0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frm_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql1, sql2 As String
Dim i, j As Integer
Dim store(1000) As Variant


Private Sub cmdDelet_Click()

sql1 = "select * from client Where Invoice_ID Like " & textInvoice_ID.Text
Set rsrecord = cnnDB.Execute(sql1)

If rsrecord.RecordCount > 0 Then


lisql = "DELETE FROM `Client` Where Invoice_ID=" & textInvoice_ID.Text

cnnDB.Execute (lisql)

lisql = "DELETE FROM `invoice` Where Invoice_ID ='" & textInvoice_ID.Text & "'"
                                                
cnnDB.Execute (lisql)

MsgBox "Deleate Success !", vbInformation, "Delete Success !"

Else

MsgBox "No Record In The Database !", vbInformation, "Record Error!"

End If

End Sub

Private Sub cmdfind_Click()
'On Error GoTo err

 '*******************Invoice Search *******************************

If Optinv.Value = True Then

sql1 = "select * from client Where Invoice_ID Like " & textInvoice_ID.Text
Set rsrecord = cnnDB.Execute(sql1)

If rsrecord.RecordCount > 0 Then

textCname.Text = rsrecord.Fields("Name")
textCaddress.Text = rsrecord.Fields("Address")
textCphone.Text = rsrecord.Fields("Phone")
textCfax.Text = rsrecord.Fields("Fax")
textCmail.Text = rsrecord.Fields("Email")
textDate.Text = rsrecord.Fields("ADate")
texttotal.Text = rsrecord.Fields("Total_Tk")

sql2 = "SELECT Particulars,P_ID,Unite_Price ,Quantity From invoice WHERE Invoice_ID like " & textInvoice_ID.Text

Set rsrecord = cnnDB.Execute(sql2)

DGProduct.Caption = "Parches By " + textCname.Text

Set DGProduct.DataSource = rsrecord

Command2.Visible = False
Command1.Visible = False


Else
MsgBox "No record on the database"


End If

End If



 '*******************Date Search *******************************



If OptionDate.Value = True Then

sql1 = "select * from client Where ADate =   #" & textDate.Text & "#"
Set rsrecord = cnnDB.Execute(sql1)

'sql1 = "select * from client Where Invoice_ID Like " & textInvoice_ID.Text
'Set rsrecord = cnnDB.Execute(sql1)

If rsrecord.RecordCount > 0 Then

textCname.Text = rsrecord.Fields("Name")
textCaddress.Text = rsrecord.Fields("Address")
textCphone.Text = rsrecord.Fields("Phone")
textCfax.Text = rsrecord.Fields("Fax")
textCmail.Text = rsrecord.Fields("Email")
textDate.Text = rsrecord.Fields("ADate")
texttotal.Text = rsrecord.Fields("Total_Tk")
textInvoice_ID.Text = rsrecord.Fields("Invoice_ID")

sql2 = "SELECT Particulars,P_ID,Unite_Price ,Quantity From invoice WHERE Invoice_ID like " & textInvoice_ID.Text

Set rsrecord2 = cnnDB.Execute(sql2)

DGProduct.Caption = "Parches By " + textCname.Text

Set DGProduct.DataSource = rsrecord2


Command1.Visible = True
Command2.Visible = True

Else
MsgBox "No record on the database"
Command1.Visible = False
Command2.Visible = False

End If

End If


' '*******************Phone Search *******************************
'
'If OptionPhone.Value = True Then
'
'sql1 = "select * from client Where Phone Like " & textCphone.Text
'Set rsrecord = cnnDB.Execute(sql1)
'
'If rsrecord.RecordCount > 0 Then
'
'textCname.Text = rsrecord.Fields("Name")
'textCaddress.Text = rsrecord.Fields("Address")
'textCphone.Text = rsrecord.Fields("Phone")
'textCfax.Text = rsrecord.Fields("Fax")
'textCmail.Text = rsrecord.Fields("Email")
'textDate.Text = rsrecord.Fields("ADate")
'texttotal.Text = rsrecord.Fields("Total_Tk")
'textInvoice_ID.Text = rsrecord.Fields("Invoice_ID")
'
''sql2 = "SELECT Particulars,P_ID,Unite_Price ,Quantity From invoice WHERE Phone like " & textCphone.Text
'sql2 = "SELECT Particulars,P_ID,Unite_Price ,Quantity From invoice WHERE Invoice_ID like " & textInvoice_ID.Text
'
'Set rsrecord2 = cnnDB.Execute(sql2)
'
'DGProduct.Caption = "Parches By " + textCname.Text
'
'Set DGProduct.DataSource = rsrecord2
'
'
'
'Else
'MsgBox "No record on the database"
'
'End If
'End If
'
'
'
'
''err:

End Sub


Private Sub cmdfind_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdfind.Font.Bold = True
End Sub

Private Sub Command1_Click()

If rsrecord.RecordCount > 0 Then

rsrecord.MoveNext

If rsrecord.EOF Then

Command1.Enabled = False
Command2.Enabled = True
MsgBox "End Of The Record !"

Else
'Command1.Enabled = False
Command2.Enabled = True

textCname.Text = rsrecord.Fields("Name")
textCaddress.Text = rsrecord.Fields("Address")
textCphone.Text = rsrecord.Fields("Phone")
textCfax.Text = rsrecord.Fields("Fax")
textCmail.Text = rsrecord.Fields("Email")
textDate.Text = rsrecord.Fields("ADate")
texttotal.Text = rsrecord.Fields("Total_Tk")
textInvoice_ID.Text = rsrecord.Fields("Invoice_ID")

sql2 = "SELECT Particulars,P_ID,Unite_Price ,Quantity From invoice WHERE Invoice_ID like " & textInvoice_ID.Text

Set rsrecord2 = cnnDB.Execute(sql2)

DGProduct.Caption = "Parches By " + textCname.Text

Set DGProduct.DataSource = rsrecord2


End If

Else
MsgBox "No Record"

End If

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command1.Font.Underline = True
 
End Sub

Private Sub Command2_Click()

If rsrecord.RecordCount > 0 Then
   
rsrecord.MovePrevious

If rsrecord.BOF Then

Command1.Enabled = True
Command2.Enabled = False
MsgBox "End Of The Record !"

Else
Command1.Enabled = True

textCname.Text = rsrecord.Fields("Name")
textCaddress.Text = rsrecord.Fields("Address")
textCphone.Text = rsrecord.Fields("Phone")
textCfax.Text = rsrecord.Fields("Fax")
textCmail.Text = rsrecord.Fields("Email")
textDate.Text = rsrecord.Fields("ADate")
texttotal.Text = rsrecord.Fields("Total_Tk")
textInvoice_ID.Text = rsrecord.Fields("Invoice_ID")


sql2 = "SELECT Particulars,P_ID,Unite_Price ,Quantity From invoice WHERE Invoice_ID like " & textInvoice_ID.Text

Set rsrecord2 = cnnDB.Execute(sql2)

DGProduct.Caption = "Parches By " + textCname.Text

Set DGProduct.DataSource = rsrecord2
End If

'MsgBox "No record on the database"

End If

End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Command2.Font.Underline = True

End Sub
Private Sub Command3_Click()

'**********************************************************

If OptionEmail.Value = True Then

sql1 = "select * from client Where Email Like " & textCmail.Text
Set rsrecord = cnnDB.Execute(sql1)

If rsrecord.RecordCount > 0 Then

textCname.Text = rsrecord.Fields("Name")
textCaddress.Text = rsrecord.Fields("Address")
textCphone.Text = rsrecord.Fields("Phone")
textCfax.Text = rsrecord.Fields("Fax")
textCmail.Text = rsrecord.Fields("Email")
textDate.Text = rsrecord.Fields("ADate")
texttotal.Text = rsrecord.Fields("Total_Tk")
textInvoice_ID.Text = rsrecord.Fields("Invoice_ID")

'sql2 = "SELECT Particulars,P_ID,Unite_Price ,Quantity From invoice WHERE Phone like " & textCphone.Text
sql2 = "SELECT Particulars,P_ID,Unite_Price ,Quantity From invoice WHERE Invoice_ID like " & textInvoice_ID.Text

Set rsrecord2 = cnnDB.Execute(sql2)

DGProduct.Caption = "Parches By " + textCname.Text

Set DGProduct.DataSource = rsrecord2



Else
MsgBox "No record on the database"

End If
End If

End Sub
Private Sub Command4_Click()

lisql = "SELECT * FROM `client`" ''Where Invoice_ID ='" & textInvoice_ID.Text & "'"

Set rsrecord = cnnDB.Execute(lisql)
         
End Sub

Private Sub Command5_Click()

'sql1 = "select * from client Where ADate =   #" & textDate.Text & "#"
'Set rsrecord = cnnDB.Execute(sql1)
'
'If rsrecord.RecordCount > 0 Then
'
'
'lisql = "DELETE FROM `Client` Where ADate =   #" & textDate.Text & "#"
'
'cnnDB.Execute (lisql)
'
'lisql = "DELETE FROM `invoice` Where Invoice_ID ='" & textInvoice_ID.Text & "'"
'
'cnnDB.Execute (lisql)
'
'MsgBox "Deleate Success !", vbInformation, "Delete Success !"
'
'Else
'
'MsgBox "No Record In The Database !", vbInformation, "Record Error!"
'
'End If
 
Dim Diffrent_Date, Date_Add As Variant
                  
'Date_Add = DateAdd("m", 4, Date)
         
rsrecord.MoveNext
         
Date_Add = rsrecord.Fields("Adate")
         
Different_Date = DateDiff("m", Date_Add, Date, vbUseSystemDayOfWeek, vbUseSystem)
                 
Text1.Text = Different_Date 'itime
                 
'        If Date = itime Then
'
'        MsgBox "Free Session End.......You Must install Registered Version", 16, "Aleart Message"
'        End


End Sub
Private Sub Command6_Click()

If rsrecord.RecordCount > 0 Then

rsrecord.MoveNext

If rsrecord.EOF Then

Command7.Enabled = False
Command6.Enabled = True
MsgBox "End Of The Record !"

Else

Text2.Text = rsrecord.Fields("Invoice_ID")
Text1.Text = rsrecord.Fields("Adate")

End If

Else
MsgBox "No Record"

End If

End Sub
Private Sub Command7_Click()

If rsrecord.RecordCount > 0 Then

rsrecord.MovePrevious

If rsrecord.BOF Then

Command7.Enabled = False
Command6.Enabled = True

MsgBox "End Of The Record !"

Else

Text2.Text = rsrecord.Fields("Invoice_ID")
Text1.Text = rsrecord.Fields("Adate")


End If

Else
MsgBox "No Record"

End If

End Sub

Private Sub Command8_Click()

Dim Diffrent_Date, Date_Add As Variant
                                                               

For i = 1 To rsrecord.RecordCount - 1

rsrecord.MoveNext
         
Date_Add = rsrecord.Fields("Adate")
         
Different_Date = DateDiff("m", Date_Add, Date, vbUseSystemDayOfWeek, vbUseSystem)
                
Text1.Text = Different_Date 'itime

If (Different_Date > 4) Then

store(j + 1) = rsrecord.Fields("Adate")

End If

Next
                          
'Date_Add = DateAdd("m", 4, Date)
         
End Sub

Private Sub Command9_Click()
j = j + 1

Text3.Text = store(j)

End Sub

Private Sub DGProduct_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdfind.Font.Bold = False
Command1.Font.Underline = False
Command2.Font.Underline = False

End Sub

Private Sub Form_Load()
'Me.Top = fMainForm.Top + 100
'Me.Left = fMainForm.Left + 100

Me.Top = frmMain.Top + frmMain.Height / 2 - 4500
Me.Left = frmMain.Left + frmMain.Width / 2 - 6200

Me.Height = 7600
Me.Width = 12000

j = 0

Optinv.Value = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdfind.Font.Bold = False
Command1.Font.Underline = False
Command2.Font.Underline = False
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdfind.Font.Bold = False
End Sub

Private Sub textDate_Change()
cmdfind.Enabled = True
End Sub

Private Sub textDate_Click()
OptionDate.Value = True

End Sub

Private Sub textInvoice_ID_Change()
Optinv.Value = True
cmdfind.Enabled = True

End Sub

Private Sub textInvoice_ID_Click()
Optinv.Value = True

End Sub
