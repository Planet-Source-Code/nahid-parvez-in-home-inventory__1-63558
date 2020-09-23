Attribute VB_Name = "mdlMain"
Option Explicit
'Main Connection
Public i, J, k As Integer ' this variable using only for the Common _Product

Public cnnDB As New ADODB.Connection
Public rsrecord As New ADODB.Recordset
'Public rsrecord2 As New ADODB.Recordset
Public lisql As String
'Public connectionstring As String
Public fMainForm As frmMain

Public Sub Main()
    Call fncMakeConnectionMainDB
    frm_user_pass.Show
End Sub

Public Sub fncMakeConnectionMainDB()

On Error Resume Next
    rsrecord.LockType = adLockPessimistic
    rsrecord.CursorType = adOpenDynamic
    cnnDB.CursorLocation = adUseClient
    
   cnnDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Records.mdb;Persist Security Info=False;Jet OLEDB:Database Password=1234"
 
   
End Sub



