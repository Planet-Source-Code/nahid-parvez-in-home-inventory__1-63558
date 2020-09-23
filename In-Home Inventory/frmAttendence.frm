VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_common_product 
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ClipControls    =   0   'False
   Icon            =   "frmAttendence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7110
   ScaleMode       =   0  'User
   ScaleWidth      =   11880
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame3 
      Caption         =   "Invoice "
      Height          =   615
      Left            =   6960
      TabIndex        =   46
      Top             =   1920
      Width           =   4815
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   50
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   255
         Left            =   1920
         TabIndex        =   49
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print Direct"
         Height          =   255
         Left            =   3390
         TabIndex        =   48
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdRefresh 
         Cancel          =   -1  'True
         Caption         =   "Refresh"
         Default         =   -1  'True
         Height          =   255
         Left            =   3720
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   1000
      End
   End
   Begin VB.TextBox textDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      CausesValidation=   0   'False
      Height          =   225
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox textInvoice_ID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      CausesValidation=   0   'False
      Height          =   225
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   6600
      Width           =   11655
      Begin VB.TextBox texttotalitem 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   70
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   177
         Width           =   615
      End
      Begin VB.CommandButton Comdundo 
         Appearance      =   0  'Flat
         Caption         =   "Undo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   36
         Top             =   177
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton ComdPrint 
         Appearance      =   0  'Flat
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   255
         Left            =   8760
         TabIndex        =   35
         ToolTipText     =   "Print Invoice"
         Top             =   177
         Width           =   1245
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   255
         Left            =   10200
         TabIndex        =   31
         ToolTipText     =   "Close"
         Top             =   177
         Width           =   1245
      End
      Begin VB.CommandButton cmdtotal 
         Caption         =   "Total "
         Height          =   255
         Left            =   7320
         TabIndex        =   27
         ToolTipText     =   "Add Total With Adjust"
         Top             =   177
         Width           =   1245
      End
      Begin VB.TextBox texttotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5400
         MaxLength       =   14
         TabIndex        =   26
         Top             =   177
         Width           =   1335
      End
      Begin VB.TextBox textadjust 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3720
         MaxLength       =   14
         TabIndex        =   25
         Top             =   177
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Total Item"
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "BDT"
         Height          =   255
         Left            =   6760
         TabIndex        =   30
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "Total"
         Height          =   255
         Left            =   4920
         TabIndex        =   29
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label20 
         Caption         =   "Adjust "
         Height          =   255
         Left            =   3240
         TabIndex        =   28
         Top             =   177
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3945
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6959
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox textCmail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8280
      TabIndex        =   18
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox textCfax 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   17
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox textCphone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   16
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox textCname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Common Accessories "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   11865
      Begin VB.TextBox textqt 
         Height          =   285
         Left            =   8760
         TabIndex        =   42
         Text            =   "Remain Q.T."
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox textqtonhand 
         Height          =   285
         Left            =   7920
         TabIndex        =   33
         Text            =   "Q.T.On"
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox Cmbtextitem 
         Height          =   315
         Left            =   4605
         TabIndex        =   32
         Text            =   "Q.T"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox textP_ID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   21
         Text            =   "P_ID"
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox comProduct 
         Height          =   315
         Left            =   5355
         TabIndex        =   8
         Text            =   "Combo2"
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox textSl_No 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Sl_No"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox textOpreator 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   9720
         TabIndex        =   3
         Text            =   "O.P.Name"
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox CombNameP 
         Height          =   315
         ItemData        =   "frmAttendence.frx":2CCA
         Left            =   1150
         List            =   "frmAttendence.frx":2CCC
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "Operator's Name"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9720
         TabIndex        =   45
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Remain Q.T."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Q.T.O.Hand"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7920
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label45 
         Caption         =   "Q.T."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4695
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Serial No"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Product ID"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3945
      Left            =   6960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2640
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6959
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   480
   End
   Begin VB.TextBox textCaddress 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      MaxLength       =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   4560
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Fax:"
      Height          =   255
      Left            =   6360
      TabIndex        =   44
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   -960
      Picture         =   "frmAttendence.frx":2CCE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12915
   End
   Begin VB.Label Label17 
      Caption         =   "Invoice ID"
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
      Left            =   5850
      TabIndex        =   39
      Top             =   2070
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "Date"
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
      Left            =   4920
      TabIndex        =   20
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Phone :"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Email :"
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label15 
      Caption         =   "Fax :"
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "frm_common_product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xItem As ListItem
Dim t1, t2, t3, t4, t5, t6, t7, t8, t9, lisql As String
Dim fcount As Integer
Dim remqt(1000) As Integer
Dim Purch_Tk(1000) As Variant






Private Sub cmdFind_function()

'**********************FIND FUNCTION***********************

lisql = "Select C.Product,C.Brand,C.P_ID from Common C Where C.Sl_No= " & textP_ID.Text
Set rsrecord = cnnDB.Execute(lisql)
Set DG.DataSource = rsrecord


End Sub

Private Sub cmdclose_Click()
Me.Hide

End Sub

Private Sub cmdclose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdclose.Font.Bold = True

End Sub



Private Sub cmdNew_Click()
i = 0
cmdSave.Enabled = False
ListView1.ListItems.Clear
texttotal.Text = 0

'****************Printer Button**********************

ComdPrint.Enabled = False

End Sub





Private Sub cmdRefresh_Click()
Dim sql2 As String

lisql = "Select C.Product,C.Brand,C.Amount_Tk,C.Quantity from Common C Where C.Sl_No= " & textSl_No.Text
sql2 = lisql + " order by P_ID"

Set rsrecord = cnnDB.Execute(sql2)
Set DG.DataSource = rsrecord


End Sub

Private Sub cmdSave_Click()

'Private
Dim sql1 As String
Dim i1, i2, i3, i4, i5, i6, i7 As Variant


 t2 = " " + Trim(textCname.Text)
 t3 = " " + Trim(textCaddress.Text)
 t4 = " " + Trim(textCphone.Text)
 t5 = " " + Trim(textCmail.Text)
 t6 = " " + Trim(textCfax.Text)
 t7 = " " + Trim(texttotal.Text)
 t8 = " " + Trim(textDate.Text)
  
' t2 = Trim(textCname.Text)
' t3 = Trim(textCaddress.Text)
' t4 = Trim(textCphone.Text)
' t5 = Trim(textCmail.Text)
' t6 = Trim(textCfax.Text)
' t7 = Trim(texttotal.Text)
' t8 = Trim(textDate.Text)
'

 lisql = "INSERT INTO client (Name,Address,Phone,Email,fax,Total_Tk,ADate)VALUES ( '" & t2 & "','" & t3 & "','" & t4 & "','" & t5 & "','" & t6 & "','" & t7 & "','" & t8 & "')"

 Set rsrecord = cnnDB.Execute(lisql)
 
 
lisql = "SELECT   * FROM client "
Set rsrecord = cnnDB.Execute(lisql)

rsrecord.MoveLast

textInvoice_ID.Text = rsrecord.Fields("Invoice_ID")
 
 
For k = 1 To Val(texttotalitem.Text)
 
 

Set xItem = ListView1.ListItems.Item(k)
    
 i1 = k - 1  'xItem.ListSubItems.Item(0)   'Serial No
 i2 = xItem.ListSubItems.Item(1)  ', , CombNameP.Text + " " + rsrecord.Fields("Product")
 i3 = xItem.ListSubItems.Item(2)  ', , rsrecord.Fields("P_ID")
 i4 = xItem.ListSubItems.Item(3) ', , Cmbtextitem.Text
 i5 = xItem.ListSubItems.Item(4) ', , rsrecord.Fields("Amount_Tk")

 i6 = Trim(textInvoice_ID.Text)
 
 i7 = Purch_Tk(k - 1) ' Store Purchase prise into the database


 'i6 = xItem.ListSubItems.Item(5) ', , Val(Cmbtextitem.Text) * rsrecord.Fields("Amount_Tk")

sql1 = "INSERT INTO invoice (Sl_No,Particulars,P_ID,Unite_Price,Invoice_ID,Quantity,Purchase)VALUES ( '" & i1 & "','" & i2 & "','" & i3 & "','" & i5 & "','" & i6 & "','" & i4 & "','" & i7 & "')"
 

Set rsrecord = cnnDB.Execute(sql1)

Next

k = 1
i = 0
 
 
'************Printer Button ********************


ComdPrint.Enabled = True

 
End Sub

Private Sub cmdtotal_Click()


texttotal.Text = Val(texttotal.Text) - Val(textadjust.Text)

i = 0



cmdtotal.Enabled = False





End Sub

Private Sub cmdtotal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdtotal.Font.Bold = True
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

'*************QUANTITY OF PRODUCT***********************

Cmbtextitem.Clear

Cmbtextitem.Text = "1" 'rsrecord.Fields("Quantity")

Cmbtextitem.ToolTipText = "Quantity On Hand " + rsrecord.Fields("Quantity") + " "

For J = 1 To rsrecord.Fields("Quantity")

Cmbtextitem.AddItem (J)

Next



'***************CALL THE REFRESH FUNCTION***************

Label4.Caption = "List of " & rsrecord.Fields("P_Name")
Call cmdRefresh_Click
Call Product_From_Common
 
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
 toolTips = count
 count = rsrecord.RecordCount
 For J = 1 To CInt(count)
 comProduct.AddItem rsrecord.Fields("Product") ' + "     On Hand " + rsrecord.Fields("Quantity")
  
 rsrecord.MoveNext
 
 Next
  
 comProduct.ToolTipText = "Total Item " & count
 
 rsrecord.MoveFirst
 comProduct.Text = rsrecord.Fields("Product") ' + "     On Hand " + rsrecord.Fields("Quantity")

End Sub



Private Sub ComdPrint_Click()
 
Dim sq, ret As String

DataReportInvoice.TopMargin = 1000

DataReportInvoice.BottomMargin = 800

ret = textInvoice_ID.Text  'InputBox("Enter the Employee id !", App.Title, 1)
  
sq = "SHAPE {SELECT * FROM `client` where Invoice_ID=" & CLng(ret) & " }  AS Command_Client APPEND ({SELECT Invoice_ID,Sl_No,Particulars,P_ID,Unite_Price,Quantity,Unite_Price*Quantity as [total] FROM `invoice`}  AS Command_Invoice RELATE 'Invoice_ID' TO 'Invoice_ID') AS Command_Invoice"

If DataEnvironment1.rsCommand_Client.State = adStateOpen Then
                    DataEnvironment1.rsCommand_Client.close
                End If

            DataEnvironment1.Commands("Command_Client").CommandText = sq
    
            With DataReportInvoice
            
            With .Sections("Section1").Controls
                    .Item("Text_QuantXUnite").DataMember = "Command_Invoice"
                    .Item("Text_QuantXUnite").DataField = "total"
                End With
            .Show
            
                        
End With









'
'Set repcon = New Class1
'
'textInvoice_ID.Text = 111
'
'lisql = "SELECT Particulars,P_ID,Quantity,Unite_Price From invoice WHERE Invoice_ID like " & textInvoice_ID.Text
'
''lisql = "select Common_Product.P_ID,Common_Product.P_Name,Common.Product from Common_Product,Common  Where Common_Product.Sl_No=" & textSl_No.Text & " "
'
'Set rsrecord = cnnDB.Execute(lisql)
'
'    With DataReport2
'
'            Set .DataSource = Nothing
'                .DataMember = ""
'            Set .DataSource = rsrecord.DataSource
'            With .Sections("Section1").Controls
'                For i = 1 To .count
'                    If TypeOf .Item(i) Is RptTextBox Then
'                        .Item(i).DataMember = ""
'                        .Item(i).DataField = rsrecord.Fields(i - 1).Name
'
'                       End If
'                Next i
'            End With
'                .Show
'        End With
'        Set rsrecord = Nothing


End Sub

Private Sub ComdPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ComdPrint.Font.Bold = True

End Sub

Private Sub Comdundo_Click()

'Set xItem = ListView1.ListItems.Remove(1, , i)

xItem.ListSubItems.Remove (i)

texttotal.Text = Val(texttotal.Text) - (Val(Cmbtextitem.Text) * rsrecord.Fields("Amount_Tk"))


xItem.ListSubItems.Clear

 i = i - 1
 
Comdundo.Enabled = False

End Sub

Private Sub Comdundo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Comdundo.Font.Bold = True

End Sub

Private Sub Command1_Click()


On Error GoTo Error

Dim sq, ret As String

DataReportInvoice.Visible = False

ret = textInvoice_ID.Text  ' InputBox("Enter the Invoice ID:", "Invoice Report", 1)

sq = "SHAPE {SELECT * FROM `client` where Invoice_ID=" & CLng(ret) & " }  AS Command_Client APPEND ({SELECT Invoice_ID,Sl_No,Particulars,P_ID,Unite_Price,Quantity,Unite_Price*Quantity as [total] FROM `invoice`}  AS Command_Invoice RELATE 'Invoice_ID' TO 'Invoice_ID') AS Command_Invoice"

DataReportInvoice.TopMargin = 1000
DataReportInvoice.BottomMargin = 600

  
If DataEnvironment1.rsCommand_Client.State = adStateOpen Then
            DataEnvironment1.rsCommand_Client.close

  End If

            DataEnvironment1.Commands("Command_Client").CommandText = sq
    
           ' Printer.Print
            With DataReportInvoice
            With .Sections("Section1").Controls
                    .Item("Text_QuantXUnite").DataMember = "Command_Invoice"
                    .Item("Text_QuantXUnite").DataField = "total"
            End With
           ' .Show
    End With
    DataReportInvoice.PrintReport (False)
 
 'Printer.EndDoc
             
Error:

End Sub









Private Sub comProduct_Click()
On Error GoTo err

Comdundo.Enabled = True

i = 0 + i

texttotalitem.Text = i + 1

Dim search As Variant


'search = Left(comProduct.Text, L1)
search = comProduct.Text

If rsrecord.State = adStateOpen Then
rsrecord.close
End If


lisql = "select * from Common where Product= '" & search & " '"
Set rsrecord = cnnDB.Execute(lisql)



rsrecord.MoveFirst


comProduct.Text = rsrecord.Fields("Product")

textqtonhand.Text = rsrecord.Fields("Quantity")

If Val(Cmbtextitem.Text) > rsrecord.Fields("Quantity") Then

MsgBox "Insufficient Stock, Please Chack Quantity", vbCritical, "Stock Under Flow"

GoTo err

End If



textqt.Text = Val(textqtonhand.Text) - Val(Cmbtextitem.Text)

'******* Using array to store remain Q.T.*****

remqt(i) = textqt.Text

Purch_Tk(i) = rsrecord.Fields("Purchase_Tk") ' Store Purchase prise into the database




Set xItem = ListView1.ListItems.Add(, , i)
 
 xItem.ListSubItems.Add 1, , CombNameP.Text + " " + rsrecord.Fields("Product")
 xItem.ListSubItems.Add 2, , rsrecord.Fields("P_ID")
 xItem.ListSubItems.Add 3, , Cmbtextitem.Text
 xItem.ListSubItems.Add 4, , rsrecord.Fields("Amount_Tk")
 xItem.ListSubItems.Add 5, , Val(Cmbtextitem.Text) * rsrecord.Fields("Amount_Tk")
 

texttotal.Text = Val(texttotal.Text) + (Val(Cmbtextitem.Text) * rsrecord.Fields("Amount_Tk"))


i = i + 1

cmdtotal.Enabled = True





err:

End Sub



Private Sub DG_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdtotal.Font.Bold = False
cmdclose.Font.Bold = False
ComdPrint.Font.Bold = False

End Sub

Private Sub Form_Load()

Dim counter As Integer
Dim i As Integer
Dim sqlload As String

'Me.Top = fMainForm .Top + 100
'Me.Left = fMainForm.Left + 100

Me.Top = frmMain.Top + frmMain.Height / 2 - 4500
Me.Left = frmMain.Left + frmMain.Width / 2 - 6200




Me.Width = 12000
Me.Height = 7600

 '***************EXECUTE QUARY FUNCTION **************


 Set rsrecord = cnnDB.Execute("Select * from Common_Product order by Sl_No asc")
  
 counter = rsrecord.RecordCount
 
 For i = 1 To CInt(counter)
 
 CombNameP.AddItem rsrecord.Fields("P_Name")
 rsrecord.MoveNext
 
 Next i
   
 rsrecord.MoveFirst

 textP_ID.Text = rsrecord.Fields("P_ID")
 textSl_No.Text = rsrecord.Fields("Sl_No")
 CombNameP.Text = rsrecord.Fields("P_Name")
 
 textOpreator.Text = frm_user_pass.Text_User_ID.Text
 
 
 
 
 
 
 
 
 '******************DATE FORMATE ***************
 
 
 textDate.Text = Format(Now, "Short Date")
 Cmbtextitem.Text = "1"
 
 '****************INITIAL STAGE*****************
 
 i = 1
 k = 1
 fcount = 0
 
 
 
ListView1.ColumnHeaders.Add , , "Sl.No.", 650           '1
ListView1.ColumnHeaders.Add , , "Particulars", 2500     '2
ListView1.ColumnHeaders.Add , , "P.ID.", 550            '3
ListView1.ColumnHeaders.Add , , "Q.T.", 550             '4
ListView1.ColumnHeaders.Add , , "Unite Price BDT", 1270 '5
ListView1.ColumnHeaders.Add , , "Amount BDT", 1200      '6

ListView1.ToolTipText = "Invoice"

 

 'ListView1.View = lvwReportl



DG.DefColWidth = 1200






 
'***************MAIN CALL FUNCTION **************

 Call CombNameP_Click
 Call cmdRefresh_Click
 Call Product_From_Common

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdtotal.Font.Bold = False
cmdclose.Font.Bold = False
ComdPrint.Font.Bold = False
End Sub




Private Sub Form_Resize()

'Me.Width = 12000
'Me.Height = 7600

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdtotal.Font.Bold = False
cmdclose.Font.Bold = False
ComdPrint.Font.Bold = False
Comdundo.Font.Bold = False
End Sub
Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdtotal.Font.Bold = False
End Sub

Private Sub textadjust_Change()
cmdtotal.Enabled = True

End Sub
Private Sub textCmail_Change()
'****************************Email Address Chack*********************************

Dim x, i As Variant
x = Len(textCmail.Text)
For i = 1 To x
If UCase(Mid(textCmail.Text, i, 1)) >= 0 And UCase(Mid(textCmail.Text, i, 1)) <= 9 Or _
Mid(textCmail.Text, i, 1) = "." Or Mid(textCmail.Text, i, 1) = "_" Or Mid(textCmail.Text, i, 1) = "-" Or Mid(textCmail.Text, i, 1) = "@" _
Or UCase(Mid(textCmail.Text, i, 1)) >= "A" And UCase(Mid(textCmail.Text, i, 1)) <= "Z" Then
Else
MsgBox "Invalid Email Address Please Correct it !", vbInformation, "Email Error !"

textCmail.Text = ""

End If
Next

End Sub

Private Sub textCmail_LostFocus()

'****************************Email Address Chack*********************************

Dim ps, at, dot, i As Variant

ps = Len(textCmail.Text)
at = 0
dot = 0
For i = 1 To ps
If Mid(textCmail.Text, i, 1) = "@" Then
at = at + 1
End If
If Mid(textCmail.Text, i, 1) = "." Then
dot = dot + 1
End If
Next
If ps > 0 And (dot = 0 Or at <> 1) Then
MsgBox "Invalid Email Address Please Correct it !", vbInformation, "Email Error !"
textCmail = ""
textCmail.SetFocus
End If

End Sub

Private Sub textCname_Change()

If textCname = "" Then
cmdSave.Enabled = False
Else
cmdSave.Enabled = True
End If

End Sub

Private Sub Timer1_Timer()
   
 Me.Caption = "   " + Format(Now, "Long Time") + ", " + Format(Now, "Long Date")
  
 DG.Caption = CombNameP.Text


End Sub


