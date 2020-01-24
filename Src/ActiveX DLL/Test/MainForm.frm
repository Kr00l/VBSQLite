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
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private hLib As Long
Private DBConnection As SQLiteConnection

Private Sub Form_Load()
' When referencing to the VBSQLite10.DLL then sqlite3win32.dll is built into it.
' Only for this test debugging the LoadLibrary is necessary.
hLib = LoadLibrary(StrPtr("sqlite3win32.dll"))
If hLib = 0 Then LoadLibrary (StrPtr(lib_dir_sqlite3win32()))
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DBConnection Is Nothing Then DBConnection.CloseDB
If hLib <> 0 Then
    FreeLibrary hLib
    hLib = 0
End If
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
On Error GoTo CATCH_EXCEPTION
With DBConnection
.Execute "INSERT INTO test_table (szText) VALUES ('" & VBA.InputBox("szText") & "')"
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
