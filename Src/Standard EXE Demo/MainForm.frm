VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "VBSQLite Demo"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "MainForm"
   ScaleHeight     =   3720
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandClose 
      Caption         =   "Close Test.db"
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   2595
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   4695
   End
   Begin VB.CommandButton CommandConnect 
      Caption         =   "Connect Test.db"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton CommandInsert 
      Caption         =   "Insert into test_table"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ISQLiteProgressHandler
Private DBConnection As SQLiteConnection

Private Sub ISQLiteProgressHandler_Callback(Cancel As Boolean)
' The SetProgressHandler method (which registers this callback) has a default value of 100 for the
' number of virtual machine instructions that are evaluated between successive invocations of this callback.
' This means that this callback is never invoked for very short running SQL statements.
DoEvents
' The operation will be interrupted if the cancel parameter is set to true.
' This can be used to implement a "cancel" button on a GUI progress dialog box.
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DBConnection Is Nothing Then DBConnection.CloseDB
End Sub

Private Sub CommandConnect_Click()
If DBConnection Is Nothing Then
    With New SQLiteConnection
    On Error Resume Next
    .OpenDB AppPath() & "Test.db", SQLiteReadWrite
    If Err.Number <> 0 Then
        Err.Clear
        If MsgBox("Test.db does not exist. Create new?", vbExclamation + vbOKCancel) <> vbCancel Then
            .OpenDB AppPath() & "Test.db", SQLiteReadWriteCreate
            .Execute "CREATE TABLE test_table (ID INTEGER PRIMARY KEY, szText TEXT)"
        End If
    End If
    On Error GoTo 0
    If .hDB <> 0 Then
        Set DBConnection = .Object
        .SetProgressHandler Me ' Registers the progress handler callback
        CommandInsert.Enabled = True
        List1.Enabled = True
        Call Requery
    End If
    End With
Else
    MsgBox "Already connected.", vbExclamation
End If
End Sub

Private Sub CommandInsert_Click()
Dim Text As String
Text = VBA.InputBox("szText")
If StrPtr(Text) = 0 Then Exit Sub
On Error GoTo CATCH_EXCEPTION
With DBConnection
.Execute "INSERT INTO test_table (szText) VALUES ('" & Text & "')"
End With
Call Requery
Exit Sub
CATCH_EXCEPTION:
MsgBox Err.Description, vbCritical + vbOKOnly
End Sub

Private Sub Requery()
On Error GoTo CATCH_EXCEPTION
List1.Clear
Dim DataSet As SQLiteDataSet
Set DataSet = DBConnection.OpenDataSet("SELECT ID, szText FROM test_table")
DataSet.MoveFirst
Do Until DataSet.EOF
    List1.AddItem DataSet!ID & "_" & DataSet!szText
    DataSet.MoveNext
Loop
Exit Sub
CATCH_EXCEPTION:
MsgBox Err.Description, vbCritical + vbOKOnly
End Sub

Private Sub CommandClose_Click()
If DBConnection Is Nothing Then
    MsgBox "Not connected.", vbExclamation
Else
    DBConnection.SetProgressHandler Nothing ' Unregisters the progress handler callback
    DBConnection.CloseDB
    Set DBConnection = Nothing
    CommandInsert.Enabled = False
    List1.Clear
    List1.Enabled = False
End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo CATCH_EXCEPTION
If List1.ListCount > 0 Then
    If KeyCode = vbKeyDelete Then
        If MsgBox("Delete?", vbQuestion + vbYesNo) <> vbNo Then
            Dim Command As SQLiteCommand
            Set Command = DBConnection.CreateCommand("DELETE FROM test_table WHERE ID = @oid")
            Command.SetParameterValue Command![@oid], Left$(List1.Text, InStr(List1.Text, "_") - 1)
            Command.Execute
            List1.RemoveItem List1.ListIndex
        End If
    End If
End If
Exit Sub
CATCH_EXCEPTION:
MsgBox Err.Description, vbCritical + vbOKOnly
End Sub
