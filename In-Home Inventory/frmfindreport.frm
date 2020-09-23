VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmfindreport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Invoice Report "
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check2"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox textID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Format          =   "dd-mm-yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Report "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
      MICON           =   "frmfindreport.frx":0000
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmfindreport.frx":001C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice ID:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmfindreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Width = 5010
Me.Height = 2055

Check1.Value = 0
Check2.Value = 0


End Sub
Private Sub LaVolpeButton1_Click()

On Error GoTo Error

Dim sq, ret, ret1 As String


'''***************************ID CHACK AND DATE CHACK *************************************


If (Check1.Value = 1) And (Check2.Value = 1) Then


If (MaskEdBox1.Text = "__/__/____" Or textID.Text = "") Then

MsgBox "Date or Invoice ID Field May Be Empty !", vbInformation, "Date Error"


Else


ret = MaskEdBox1.Text

ret1 = textID.Text

sq = "SHAPE {SELECT * FROM `client` where  Adate=#" & ret & "# AND Invoice_ID=" & ret1 & " }  AS Command_Client APPEND ({SELECT Invoice_ID,Sl_No,Particulars,P_ID,Unite_Price,Quantity,Unite_Price*Quantity as [total] FROM `invoice`}  AS Command_Invoice RELATE 'Invoice_ID' TO 'Invoice_ID') AS Command_Invoice"

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

 End If
End If



''********************************* SEARCH BY DATE  *********************************


If Check1.Value = 1 And Check2.Value = 0 Then


If MaskEdBox1.Text = "__/__/____" Then

MsgBox "Date Field May Be Empty !", vbInformation, "Date Error"

Else


ret = MaskEdBox1.Text

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


End If

End If


''***************************ID CHACK  ********************************************


If Check2.Value = 1 And Check1.Value = 0 Then


If textID.Text = "" Then

MsgBox "Invoice ID Field May Be Empty !", vbInformation, "Date Error"

Else



ret = textID.Text
sq = "SHAPE {SELECT * FROM `client` where  Invoice_ID=" & ret & "}  AS Command_Client APPEND ({SELECT Invoice_ID,Sl_No,Particulars,P_ID,Unite_Price,Quantity,Unite_Price*Quantity as [total] FROM `invoice`}  AS Command_Invoice RELATE 'Invoice_ID' TO 'Invoice_ID') AS Command_Invoice"


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


 End If

End If




Error:

End Sub



Private Sub MaskEdBox1_Change()

Check1.Value = 1

End Sub

Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    
    Call LaVolpeButton1_Click
    
End If
End Sub

Private Sub textID_Change()

Check2.Value = 1

End Sub

Private Sub textID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    
    Call LaVolpeButton1_Click
    
End If

End Sub
