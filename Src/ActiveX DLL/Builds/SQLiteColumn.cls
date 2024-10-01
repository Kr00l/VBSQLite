VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLiteColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If
Private ShadowObjectPointer_0 As LongPtr, ShadowObjectPointer_1 As LongPtr
Private PropIndex As Long, PropKey As String
Private PropName As String
Private PropTypeName As String
Private PropOriginalName As String
Private PropTableName As String
Private PropDBName As String

Private Function ShadowDataSet() As SQLiteDataSet
Set ShadowDataSet = PtrToObj(ShadowObjectPointer_0)
End Function

Private Function ShadowCursor() As SQLiteCursor
Set ShadowCursor = PtrToObj(ShadowObjectPointer_1)
End Function

#If VBA7 Then
Friend Sub FInit(ByVal ObjectPointer_0 As LongPtr, ByVal ObjectPointer_1 As LongPtr, ByVal Index As Long, ByRef Key As String, ByVal hStmt As LongPtr, ByVal iCol As Long)
#Else
Friend Sub FInit(ByVal ObjectPointer_0 As Long, ByVal ObjectPointer_1 As Long, ByVal Index As Long, ByRef Key As String, ByVal hStmt As Long, ByVal iCol As Long)
#End If
If hStmt <> NULL_PTR And iCol > -1 Then
    PropName = SQLiteUTF8PtrToStr(stub_sqlite3_column_name(hStmt, iCol))
    PropTypeName = SQLiteUTF8PtrToStr(stub_sqlite3_column_decltype(hStmt, iCol))
    PropOriginalName = SQLiteUTF8PtrToStr(stub_sqlite3_column_origin_name(hStmt, iCol))
    PropTableName = SQLiteUTF8PtrToStr(stub_sqlite3_column_table_name(hStmt, iCol))
    PropDBName = SQLiteUTF8PtrToStr(stub_sqlite3_column_database_name(hStmt, iCol))
    Key = PropName
Else
    ShadowObjectPointer_0 = ObjectPointer_0
    ShadowObjectPointer_1 = ObjectPointer_1
    PropIndex = Index
    PropKey = Key
End If
End Sub

Public Property Get Index() As Long
Attribute Index.VB_Description = "Returns the index in the collection of the column item."
Index = PropIndex
End Property

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key used in the collection."
Key = PropKey
End Property

Public Property Get Value() As Variant
Attribute Value.VB_Description = "Returns the value of the column item from the current record."
Attribute Value.VB_UserMemId = 0
If ShadowObjectPointer_0 <> NULL_PTR Then
    Dim Row As Long
    Row = ShadowDataSet.Position - 1
    If Row > -1 Then
        Value = ShadowDataSet.ValueMatrix(Row, PropIndex - 1)
    Else
        Err.Raise Number:=vbObjectError, Description:="Either BOF or EOF is True"
    End If
Else
    Value = ShadowCursor.Value(PropIndex - 1)
End If
End Property

Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the name of the column item."
Name = PropName
End Property

Public Property Get TypeName() As String
Attribute TypeName.VB_Description = "Returns the declared datatype of the column item."
TypeName = PropTypeName
End Property

Public Property Get OriginalName() As String
Attribute OriginalName.VB_Description = "Returns the original un-aliased name of the column item."
OriginalName = PropOriginalName
End Property

Public Property Get TableName() As String
Attribute TableName.VB_Description = "Returns the originating table name of the column item."
TableName = PropTableName
End Property

Public Property Get DBName() As String
Attribute DBName.VB_Description = "Returns the originating database name of the column item."
DBName = PropDBName
End Property
