VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H80000004&
   Caption         =   "In-Home Automation"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11400
   Icon            =   "frmMain.frx":0000
   Picture         =   "frmMain.frx":2CCA
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1800
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E3EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E547
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F399
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FC75
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":112D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13FAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14E01
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1645D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19139
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BE15
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D471
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DD4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EB9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21879
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8040
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   476
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14473
            Text            =   "In-Home  Soft System"
            TextSave        =   "In-Home  Soft System"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "7/23/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:50 AM"
         EndProperty
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
   End
   Begin VB.Menu File_etc 
      Caption         =   "&System Works"
      Begin VB.Menu SideBar 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Sys Task|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:&H00FFC0C0&|Gradient}"
      End
      Begin VB.Menu new 
         Caption         =   "{IMG:4}New"
         Shortcut        =   ^N
      End
      Begin VB.Menu save 
         Caption         =   "{IMG:4}Save "
         Shortcut        =   ^S
      End
      Begin VB.Menu CPass 
         Caption         =   "{IMG:10}Change Password"
         Shortcut        =   {F6}
      End
      Begin VB.Menu UserID 
         Caption         =   "{IMG:6}Change User ID"
         Shortcut        =   {F7}
      End
      Begin VB.Menu Logout 
         Caption         =   "{IMG:2}Logout...."
         Shortcut        =   {F2}
      End
      Begin VB.Menu Print 
         Caption         =   "{IMG:14}Print Invoice"
      End
      Begin VB.Menu close 
         Caption         =   "{IMG:13} In-Home Info "
         Shortcut        =   ^I
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "{IMG:1} Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Action Work "
      Begin VB.Menu mnu1 
         Caption         =   "-Company Task"
      End
      Begin VB.Menu mnuEA 
         Caption         =   "{IMG:3}&Common Product"
         Shortcut        =   ^P
      End
      Begin VB.Menu client 
         Caption         =   "{IMG:12}&Client Information"
      End
      Begin VB.Menu sep4 
         Caption         =   "-Admin Works"
      End
      Begin VB.Menu Upgrade 
         Caption         =   "{IMG:7}&Upgrade Information"
         Shortcut        =   ^U
      End
      Begin VB.Menu New_User 
         Caption         =   "{IMG:12}Create &New User "
      End
      Begin VB.Menu User_Detail 
         Caption         =   "{IMG:12}User Detail"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCon 
         Caption         =   "{IMG:6}&Emp Information"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Report && Print"
      Begin VB.Menu Report 
         Caption         =   "-All Reports"
      End
      Begin VB.Menu mnu4 
         Caption         =   "{IMG:5} Full Stock Report"
      End
      Begin VB.Menu IStockrpt 
         Caption         =   "{IMG:4} Individual Stock Report"
      End
      Begin VB.Menu Invoice 
         Caption         =   "{IMG:4}Find Invoice Report"
      End
      Begin VB.Menu Monthly_Report 
         Caption         =   "{IMG:4} Daily Report"
      End
      Begin VB.Menu Yearly_Report 
         Caption         =   "{IMG:4}Yearly Report"
      End
      Begin VB.Menu Profit_Report 
         Caption         =   "{IMG:4}Profit Report"
      End
      Begin VB.Menu Admin 
         Caption         =   "{IMG:4}Admin && Stuff"
      End
      Begin VB.Menu mnure 
         Caption         =   "{IMG:4}Client Report"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu det 
         Caption         =   "Details"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu Help 
         Caption         =   "{IMG:8} Help ?"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "{IMG:9} &About !"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Admin_Click()
'DataReport_employee.TopMargin = 1000
'
'DataReport_employee.BottomMargin = 800

DataReport_employee.Show

End Sub

Private Sub client_Click()
frm_Client.Show
End Sub
Private Sub close_Click()
frm_In_homeinfo.Show
End Sub

Private Sub CPass_Click()

lisql = "SELECT * FROM `Identity` WHERE User_ID ='" & frm_user_pass.Text_User_ID.Text & "' "

Set rsrecord = cnnDB.Execute(lisql)

frm_Change_userid_passwd.Text_User_ID.Text = rsrecord.Fields("User_ID")
frm_Change_userid_passwd.Text1.Text = rsrecord.Fields("Passwd")

frm_Change_userid_passwd.Height = 2130
frm_Change_userid_passwd.Caption = "Change User Password "
frm_Change_userid_passwd.LaVolpeButton1.Caption = "OK"
frm_Change_userid_passwd.Text_User_ID.Enabled = False
frm_Change_userid_passwd.Text_FName.Visible = False
frm_Change_userid_passwd.Label5.Visible = False



frm_Change_userid_passwd.Show

End Sub

Private Sub det_Click()

''***************************ID CHACK  ********************************************
''
''Dim sq, ret, ret1 As String
''
'''DataReport_Profit
''
''ret = 144
''
''sq = "SHAPE {SELECT * FROM `client` where  Invoice_ID=" & ret & "}  AS Command_Client APPEND ({SELECT Invoice_ID,Sl_No,Particulars,P_ID,Unite_Price,Quantity,Purchase,Quantity*Purchase as [Purchase_Tk],Unite_Price*Quantity as [total] FROM `invoice`}  AS Command_Invoice RELATE 'Invoice_ID' TO 'Invoice_ID') AS Command_Invoice"
'''sq = "SHAPE {SELECT * FROM `client` where  Invoice_ID=" & ret & "}  AS Command_Client APPEND ({SELECT Invoice_ID,Sl_No,Particulars,P_ID,Unite_Price,Quantity,Purchase,Unite_Price*Quantity as [total] FROM `invoice`}  AS Command_Invoice RELATE 'Invoice_ID' TO 'Invoice_ID') AS Command_Invoice"
''
''
''DataReport_Profit.TopMargin = 1000
''DataReport_Profit.BottomMargin = 600
''
''If DataEnvironment1.rsCommand_Client.State = adStateOpen Then
''   DataEnvironment1.rsCommand_Client.close
''   End If
''
''        DataEnvironment1.Commands("Command_Client").CommandText = sq
''
''        With DataReport_Profit ' DataReport_Daily_Sale
''
''            With .Sections("Section1").Controls
''                    .Item("Text_QuantXUnite").DataMember = "Command_Invoice"
''                    .Item("Text_QuantXUnite").DataField = "total"
''
''                    .Item("Text_Purchase").DataMember = "Command_Invoice"
''                    .Item("Text_Purchase").DataField = "Purchase_Tk"
''
''                End With
''
''
''            With .Sections("Section7").Controls
''                    .Item("Text_Purchase_Total").FunctionType = 0
''                    .Item("Text_Purchase_Total").DataMember = "Command_Invoice"
''                    .Item("Text_Purchase_Total").DataField = "Purchase_Tk"
''
''                End With
''
''
''
''            .Show
''
''    End With
''
''
''
''
''
End Sub

Private Sub exit_Click()
Unload Me
Unload frm_user_pass
Unload About
cnnDB.close
End Sub

Private Sub Invoice_Click()

frmfindreport.Show

'On Error GoTo Error
'
'Dim sq, ret As String
'
'
'ret = InputBox("Enter the Invoice ID:", "Invoice Report", 1)
'
'
'
'sq = "SHAPE {SELECT * FROM `client` where Invoice_ID=" & CLng(ret) & " }  AS Command_Client APPEND ({SELECT Invoice_ID,Sl_No,Particulars,P_ID,Unite_Price,Quantity,Unite_Price*Quantity as [total] FROM `invoice`}  AS Command_Invoice RELATE 'Invoice_ID' TO 'Invoice_ID') AS Command_Invoice"
'
'DataReportInvoice.TopMargin = 1000
'DataReportInvoice.BottomMargin = 600
'
'
'If DataEnvironment1.rsCommand_Client.State = adStateOpen Then
'            DataEnvironment1.rsCommand_Client.close
'
'  End If
'
'            DataEnvironment1.Commands("Command_Client").CommandText = sq
'
'            With DataReportInvoice
'
'
'            With .Sections("Section1").Controls
'                    .Item("Text_QuantXUnite").DataMember = "Command_Invoice"
'                    .Item("Text_QuantXUnite").DataField = "total"
'
'                End With
'            .Show
'    End With
'
'Error:

End Sub

Private Sub IStockrpt_Click()

On Error GoTo Error


Dim sq, ret As String


ret = InputBox("Enter the Product No # : ", "Invoice Report", 1)



sq = "SHAPE {SELECT * FROM `Common_Product` Where Sl_No= " & CLng(ret) & " }  AS Command_Common_Product APPEND ({SELECT * FROM `Common`}  AS Command_Common RELATE 'Sl_No' TO 'Sl_No') AS Command_Common"

DataReport_Stock_list.TopMargin = 1000
DataReport_Stock_list.BottomMargin = 600
DataReport_Stock_list.LeftMargin = 1200
DataReport_Stock_list.RightMargin = 1000

If DataEnvironment1.rsCommand_Common_Product.State = adStateOpen Then
                    DataEnvironment1.rsCommand_Common_Product.close
            End If

            DataEnvironment1.Commands("Command_Common_Product").CommandText = sq

            With DataReport_Stock_list
            .Show

            End With

Error:

End Sub



Private Sub Logout_Click()
Me.Hide

frm_user_pass.Show

End Sub

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
  
  Call mnuEA_Click
   
 '********************To CREATE BAR MANUE******************************
   
   SetMenus hwnd, ImageList

End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    

    
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
 Unload Me
 Unload frm_user_pass
 Unload About


 
 '******************Connection Close******************************
 
 'cnnDB.close
 
    
End Sub

Private Sub mnu4_Click()

On Error GoTo Error

DataReport_Stock_list.TopMargin = 1000
DataReport_Stock_list.BottomMargin = 600
DataReport_Stock_list.LeftMargin = 1200
DataReport_Stock_list.RightMargin = 1000


sq = "SHAPE {SELECT * FROM `Common_Product`}  AS Command_Common_Product APPEND ({SELECT * FROM `Common`}  AS Command_Common RELATE 'Sl_No' TO 'Sl_No') AS Command_Common"


If DataEnvironment1.rsCommand_Common_Product.State = adStateOpen Then
                    DataEnvironment1.rsCommand_Common_Product.close
            End If

            DataEnvironment1.Commands("Command_Common_Product").CommandText = sq
    
            With DataReport_Stock_list
            .Show
           
            End With



'DataReport_Stock_list.Show

Error:

End Sub

Private Sub mnuCon_Click()
'frmEmpInfo.Show
frm_Employee.Show
End Sub

Private Sub mnuEA_Click()
frm_common_product.Show
End Sub
Private Sub mnuHelpAbout_Click()
About.Show
End Sub

Private Sub Monthly_Report_Click()

On Error GoTo Error

Dim sq, ret As String


ret = Date 'InputBox("Enter the Date DD\MM\YY:", "Daily Sale Report", 1)


sq = "SHAPE {SELECT * FROM `client` where  Adate=#" & ret & "#}  AS Command_Client APPEND ({SELECT Invoice_ID,Sl_No,Particulars,P_ID,Unite_Price,Quantity,Unite_Price*Quantity as [total] FROM `invoice`}  AS Command_Invoice RELATE 'Invoice_ID' TO 'Invoice_ID') AS Command_Invoice"


DataReport_Daily_Sale.TopMargin = 1000
DataReport_Daily_Sale.BottomMargin = 600

If DataEnvironment1.rsCommand_Client.State = adStateOpen Then
                    DataEnvironment1.rsCommand_Client.close
                  
 End If
            DataEnvironment1.Commands("Command_Client").CommandText = sq
            With DataReport_Daily_Sale
            With .Sections("Section1").Controls
                    .Item("Text_QuantXUnite").DataMember = "Command_Invoice"
                    .Item("Text_QuantXUnite").DataField = "total"
                End With
            .Show

    End With
Error:


'frmfindreport.Show

End Sub

Private Sub new_Click()
i = 0
frm_common_product.cmdSave.Enabled = False
frm_common_product.ListView1.ListItems.Clear
frm_common_product.texttotal.Text = 0

'****************Printer Button**********************

frm_common_product.ComdPrint.Enabled = False

End Sub

Private Sub New_User_Click()

frm_Change_userid_passwd.Caption = "Add New User"
frm_Change_userid_passwd.LaVolpeButton1.Caption = "Add New User"

frm_Change_userid_passwd.Check1.Visible = False
frm_Change_userid_passwd.Show


End Sub
Private Sub Profit_Report_Click()
frmProfit.Show
End Sub

Private Sub Upgrade_Click()
frmUpgrade.Show
End Sub

Private Sub User_Detail_Click()

lisql = "SELECT * FROM `Identity` " ' WHERE User_ID ='" & frm_user_pass.Text_User_ID.Text & "' "

Set rsrecord = cnnDB.Execute(lisql)

frm_Change_userid_passwd.Text_User_ID.Text = rsrecord.Fields("User_ID")
frm_Change_userid_passwd.Text1.Text = rsrecord.Fields("Passwd")
frm_Change_userid_passwd.TextCdate.Text = rsrecord.Fields("Cdate")
frm_Change_userid_passwd.TextMdate.Text = rsrecord.Fields("Mdate")
frm_Change_userid_passwd.TextComants.Text = rsrecord.Fields("Comants")
frm_Change_userid_passwd.Text_FName.Text = rsrecord.Fields("FName")

frm_Change_userid_passwd.Height = 3255
frm_Change_userid_passwd.Caption = "User Details "
frm_Change_userid_passwd.LaVolpeButton1.Visible = False
frm_Change_userid_passwd.Text_User_ID.Enabled = False
frm_Change_userid_passwd.Text1.Enabled = False
frm_Change_userid_passwd.TextCdate.Visible = True
frm_Change_userid_passwd.TextMdate.Visible = True
frm_Change_userid_passwd.Label2.Visible = True
frm_Change_userid_passwd.Label3.Visible = True
frm_Change_userid_passwd.Label4.Visible = True
frm_Change_userid_passwd.LaVolpeButtonNext.Visible = True
frm_Change_userid_passwd.LaVolpeButtonPrev.Visible = True
frm_Change_userid_passwd.TextComants.Visible = True

frm_Change_userid_passwd.Check1.Left = 2500
frm_Change_userid_passwd.Check1.Top = 980
frm_Change_userid_passwd.Check1.Width = 250

'frm_Change_userid_passwd.Check1.Visible = False
frm_Change_userid_passwd.CheckAdministator.Visible = False
frm_Change_userid_passwd.Check2.Visible = False
frm_Change_userid_passwd.Check3.Visible = False
frm_Change_userid_passwd.Label6.Visible = False

'frm_Change_userid_passwd.Show

frm_Change_userid_passwd.Show

End Sub

Private Sub UserID_Click()


lisql = "SELECT * FROM `Identity` WHERE User_ID ='" & frm_user_pass.Text_User_ID.Text & "' "

Set rsrecord = cnnDB.Execute(lisql)

frm_Change_userid_passwd.Text_User_ID.Text = rsrecord.Fields("User_ID")
frm_Change_userid_passwd.Text1.Text = rsrecord.Fields("Passwd")


frm_Change_userid_passwd.Height = 2130
frm_Change_userid_passwd.Caption = "Change User ID "
frm_Change_userid_passwd.LaVolpeButton1.Caption = "Change"
frm_Change_userid_passwd.Text1.Enabled = False
frm_Change_userid_passwd.Text_FName.Visible = False
frm_Change_userid_passwd.Label5.Visible = False

frm_Change_userid_passwd.Show

End Sub
