VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLiteColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private PropColumn As Collection

Private Sub Class_Initialize()
Set PropColumn = New Collection
End Sub

Private Sub Class_Terminate()
Set PropColumn = Nothing
End Sub

Public Function NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = PropColumn.[_NewEnum]
End Function

Friend Sub FInit(ByVal Object As SQLiteDataSet, ByVal hStmt As Long, ByRef Count As Long)
Count = stub_sqlite3_column_count(hStmt)
If Count > 0 Then
    Dim ObjectPointer As Long
    ObjectPointer = ObjPtr(Object)
    Dim iCol As Long, Item As SQLiteColumn, ErrValue As Long
    Dim Key As String, Key_Ext As String, Key_Cnt As Long
    For iCol = 0 To Count - 1
        Set Item = New SQLiteColumn
        Key = vbNullString
        Key_Ext = vbNullString
        Key_Cnt = 0
        Item.FInit 0, 0, Key, hStmt, iCol
        On Error Resume Next
        PropColumn.Add Item, Key
        ErrValue = Err.Number
        On Error GoTo 0
        If ErrValue <> 0 Then
            Key_Ext = Key & "_2"
            Key_Cnt = 2
            While Me.Exists(Key_Ext) = True
                Key_Cnt = Key_Cnt + 1
                Key_Ext = Key & "_" & Key_Cnt
            Wend
            Key = Key_Ext
            PropColumn.Add Item, Key
        End If
        Item.FInit ObjectPointer, iCol + 1, Key, 0, -1
    Next iCol
End If
End Sub

Public Property Get Item(ByVal Index As Variant) As SQLiteColumn
Attribute Item.VB_Description = "Returns a column item given its index or key."
Attribute Item.VB_UserMemId = 0
Set Item = PropColumn(Index)
End Property

Public Function Exists(ByVal Index As Variant) As Boolean
Attribute Exists.VB_Description = "Returns a value indicating if a given index or key in the column items collection exists."
On Error Resume Next
PropColumn.Item Index
Exists = CBool(Err.Number = 0)
On Error GoTo 0
End Function

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of column items in the collection."
Count = PropColumn.Count
End Property
