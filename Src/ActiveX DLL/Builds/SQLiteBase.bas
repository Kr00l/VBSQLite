Attribute VB_Name = "SQLiteBase"
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
#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
#End If
Private Const CP_UTF8 As Long = 65001
Private Const JULIANDAY_OFFSET As Double = 2415018.5
Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RELEASE As Long = &H8000&
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private SQLiteRefCount As Long

Public Sub SQLiteAddRef()
' It is recommended that applications always invoke sqlite3_initialize() directly prior to using any other functions.
' Future releases of SQLite may require this. In other words, the behavior exhibited when SQLite is compiled with SQLITE_OMIT_AUTOINIT might become the default behavior in some future release of SQLite.
If SQLiteRefCount = 0 Then stub_sqlite3_initialize
SQLiteRefCount = SQLiteRefCount + 1
End Sub

Public Sub SQLiteRelease()
SQLiteRefCount = SQLiteRefCount - 1
If SQLiteRefCount = 0 Then stub_sqlite3_shutdown
End Sub

#If VBA7 Then
Public Sub SQLiteOverloadBuiltinFunctions(ByVal hDB As LongPtr)
#Else
Public Sub SQLiteOverloadBuiltinFunctions(ByVal hDB As Long)
#End If
Const STR_LOWER_UTF8 As Currency = 49132859.7868@
Const STR_UPPER_UTF8 As Currency = 49132813.9381@
Const STR_LIKE_UTF8 As Currency = 170153.8156@
Const STR_NOCASE_UTF8 As Currency = 11154622955.0958@
If hDB <> NULL_PTR Then
    stub_sqlite3_create_function_v2 hDB, VarPtr(STR_LOWER_UTF8), 1, SQLITE_UTF8 Or SQLITE_DETERMINISTIC Or SQLITE_INNOCUOUS, 0, AddressOf SQLiteFunctionLowerUpper, NULL_PTR, NULL_PTR, NULL_PTR
    stub_sqlite3_create_function_v2 hDB, VarPtr(STR_LOWER_UTF8), 1, SQLITE_UTF16 Or SQLITE_DETERMINISTIC Or SQLITE_INNOCUOUS, 0, AddressOf SQLiteFunctionLowerUpper, NULL_PTR, NULL_PTR, NULL_PTR
    stub_sqlite3_create_function_v2 hDB, VarPtr(STR_UPPER_UTF8), 1, SQLITE_UTF8 Or SQLITE_DETERMINISTIC Or SQLITE_INNOCUOUS, 1, AddressOf SQLiteFunctionLowerUpper, NULL_PTR, NULL_PTR, NULL_PTR
    stub_sqlite3_create_function_v2 hDB, VarPtr(STR_UPPER_UTF8), 1, SQLITE_UTF16 Or SQLITE_DETERMINISTIC Or SQLITE_INNOCUOUS, 1, AddressOf SQLiteFunctionLowerUpper, NULL_PTR, NULL_PTR, NULL_PTR
    stub_sqlite3_create_function_v2 hDB, VarPtr(STR_LIKE_UTF8), 2, SQLITE_UTF8 Or SQLITE_DETERMINISTIC Or SQLITE_INNOCUOUS, 0, AddressOf SQLiteFunctionLike, NULL_PTR, NULL_PTR, NULL_PTR
    stub_sqlite3_create_function_v2 hDB, VarPtr(STR_LIKE_UTF8), 3, SQLITE_UTF8, 0, NULL_PTR, NULL_PTR, NULL_PTR, NULL_PTR
    stub_sqlite3_create_collation_v2 hDB, VarPtr(STR_NOCASE_UTF8), SQLITE_UTF8, 0, AddressOf SQLiteFunctionNoCaseCollating, NULL_PTR
End If
End Sub

#If VBA7 Then
Public Function SQLiteFunctionLowerUpper CDecl(ByVal pCtx As LongPtr, ByVal cArg As Long, ByVal pArgValue As LongPtr) As Long
#Else
Public Function SQLiteFunctionLowerUpper(ByVal pCtx As Long, ByVal cArg As Long, ByVal pArgValue As Long) As Long
#End If
If cArg = 1 Then
    Dim pValue As LongPtr
    CopyMemory pValue, ByVal pArgValue, PTR_SIZE
    Dim Text As String, cbText As Long
    cbText = stub_sqlite3_value_bytes16(pValue)
    If cbText > 0 Then
        If stub_sqlite3_user_data(pCtx) = 0 Then
            Text = LCase$(SQLiteUTF16PtrToStr(stub_sqlite3_value_text16(pValue), cbText \ 2))
        Else
            Text = UCase$(SQLiteUTF16PtrToStr(stub_sqlite3_value_text16(pValue), cbText \ 2))
        End If
        stub_sqlite3_result_text16 pCtx, StrPtr(Text), -1, SQLITE_TRANSIENT
    Else
        stub_sqlite3_result_text16 pCtx, NULL_PTR, 0, SQLITE_STATIC
    End If
End If
End Function

#If VBA7 Then
Public Function SQLiteFunctionLike CDecl(ByVal pCtx As LongPtr, ByVal cArg As Long, ByVal pArgValue As LongPtr) As Long
#Else
Public Function SQLiteFunctionLike(ByVal pCtx As Long, ByVal cArg As Long, ByVal pArgValue As Long) As Long
#End If
If cArg = 2 Then
    Dim pValue(0 To 1) As LongPtr
    CopyMemory pValue(0), ByVal pArgValue, PTR_SIZE * 2
    If stub_sqlite3_value_bytes(pValue(0)) > SQLITE_MAX_LIKE_PATTERN_LENGTH Then
        stub_sqlite3_result_error_toobig pCtx
        Exit Function
    End If
    Dim szPattern As String, szString As String
    szPattern = SQLiteUTF8PtrToStr(stub_sqlite3_value_text(pValue(0)), stub_sqlite3_value_bytes(pValue(0)))
    szString = SQLiteUTF8PtrToStr(stub_sqlite3_value_text(pValue(1)), stub_sqlite3_value_bytes(pValue(1)))
    If TextCompareLike(szString, szPattern) Then
        stub_sqlite3_result_int pCtx, 1
    Else
        stub_sqlite3_result_int pCtx, 0
    End If
End If
End Function

#If VBA7 Then
Public Function SQLiteFunctionNoCaseCollating CDecl(ByVal pNotUsed As LongPtr, ByVal nKey1 As Long, ByVal pKey1 As LongPtr, ByVal nKey2 As Long, ByVal pKey2 As LongPtr) As Long
#Else
Public Function SQLiteFunctionNoCaseCollating(ByVal pNotUsed As Long, ByVal nKey1 As Long, ByVal pKey1 As Long, ByVal nKey2 As Long, ByVal pKey2 As Long) As Long
#End If
SQLiteFunctionNoCaseCollating = StrComp(SQLiteUTF8PtrToStr(pKey1, nKey1), SQLiteUTF8PtrToStr(pKey2, nKey2), vbTextCompare)
End Function

#If VBA7 Then
Public Function SQLiteProgressHandlerCallback CDecl(ByVal pArg As ISQLiteProgressHandler) As Long
#Else
Public Function SQLiteProgressHandlerCallback(ByVal pArg As ISQLiteProgressHandler) As Long
#End If
Dim Cancel As Boolean
pArg.Callback Cancel
If Cancel = False Then SQLiteProgressHandlerCallback = 0 Else SQLiteProgressHandlerCallback = 1
End Function

#If VBA7 Then
Public Function SQLiteBlobToByteArray(ByVal Ptr As LongPtr, ByVal Size As Long) As Variant
#Else
Public Function SQLiteBlobToByteArray(ByVal Ptr As Long, ByVal Size As Long) As Variant
#End If
If Ptr <> NULL_PTR And Size > 0 Then
    Dim B() As Byte
    ReDim B(0 To (Size - 1)) As Byte
    CopyMemory B(0), ByVal Ptr, Size
    SQLiteBlobToByteArray = B()
Else
    SQLiteBlobToByteArray = Null
End If
End Function

#If VBA7 Then
Public Function SQLiteUTF8PtrToStr(ByVal Ptr As LongPtr, Optional ByVal Size As Long = -1) As String
#Else
Public Function SQLiteUTF8PtrToStr(ByVal Ptr As Long, Optional ByVal Size As Long = -1) As String
#End If
If Ptr <> NULL_PTR Then
    If Size = -1 Then Size = lstrlenA(Ptr)
    If Size > 0 Then
        Dim Length As Long
        Length = MultiByteToWideChar(CP_UTF8, 0, Ptr, Size, NULL_PTR, 0)
        If Length > 0 Then
            SQLiteUTF8PtrToStr = Space$(Length)
            MultiByteToWideChar CP_UTF8, 0, Ptr, Size, StrPtr(SQLiteUTF8PtrToStr), Length
        End If
    End If
End If
End Function

#If VBA7 Then
Public Function SQLiteUTF16PtrToStr(ByVal Ptr As LongPtr, Optional ByVal Size As Long = -1) As String
#Else
Public Function SQLiteUTF16PtrToStr(ByVal Ptr As Long, Optional ByVal Size As Long = -1) As String
#End If
If Ptr <> NULL_PTR Then
    If Size = -1 Then Size = lstrlen(Ptr)
    If Size > 0 Then
        SQLiteUTF16PtrToStr = Space$(Size)
        CopyMemory ByVal StrPtr(SQLiteUTF16PtrToStr), ByVal Ptr, Size * 2
    End If
End If
End Function

Public Function CDateToJulianDay(ByVal DateValue As Date) As Double
If CDbl(DateValue) >= 0 Then
    CDateToJulianDay = CDbl(DateValue) + JULIANDAY_OFFSET
Else
    Dim Temp As Double
    Temp = -Int(-CDbl(DateValue))
    CDateToJulianDay = Temp - (CDbl(DateValue) - Temp) + JULIANDAY_OFFSET
End If
End Function

Public Function CJulianDayToDate(ByVal JulianDay As Double) As Date
Const MIN_DATE As Double = -657434# + JULIANDAY_OFFSET ' 01/01/0100
Const MAX_DATE As Double = 2958465# + JULIANDAY_OFFSET ' 12/31/9999
If JulianDay < MIN_DATE Or JulianDay > MAX_DATE Then Exit Function
If JulianDay >= JULIANDAY_OFFSET Then
    CJulianDayToDate = CDate(JulianDay - JULIANDAY_OFFSET)
Else
    Dim DateDbl As Double, Temp As Double
    DateDbl = JulianDay - JULIANDAY_OFFSET
    Temp = Int(DateDbl)
    CJulianDayToDate = CDate(Temp + (Temp - DateDbl))
End If
End Function
