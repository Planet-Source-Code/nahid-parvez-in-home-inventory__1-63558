VERSION 5.00
Begin VB.Form frmEmpInfo 
   Caption         =   "Eployee Information"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Caption         =   "Provide On Demand"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "frmEmpInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()


Me.Height = 3255
Me.Width = 4845

End Sub
