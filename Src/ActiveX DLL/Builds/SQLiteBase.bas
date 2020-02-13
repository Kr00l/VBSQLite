Attribute VB_Name = "SQLiteBase"
Option Explicit
Option Compare Text
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByRef lpAddress As Long, ByVal dwSize As Long, ByVal flAllocType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByRef lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 As Long = 65001
Private Const JULIANDAY_OFFSET As Double = 2415018.5
Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RELEASE As Long = &H8000&
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private SQLiteRefCount As Long
Private SQLiteCDeclCallbackLowerUpper As Long
Private SQLiteCDeclCallbackLike As Long
Private SQLiteCDeclCallbackNoCaseCollating As Long

Public Sub SQLiteAddRef()
' It is recommended that applications always invoke sqlite3_initialize() directly prior to using any other functions.
' Future releases of SQLite may require this. In other words, the behavior exhibited when SQLite is compiled with SQLITE_OMIT_AUTOINIT might become the default behavior in some future release of SQLite.
If SQLiteRefCount = 0 Then
    stub_sqlite3_initialize
    If SQLiteCDeclCallbackLowerUpper = 0 Then SQLiteCDeclCallbackLowerUpper = SQLiteCreateCDeclCallback(AddressOf SQLiteFunctionLowerUpper, 12)
    If SQLiteCDeclCallbackLike = 0 Then SQLiteCDeclCallbackLike = SQLiteCreateCDeclCallback(AddressOf SQLiteFunctionLike, 12)
    If SQLiteCDeclCallbackNoCaseCollating = 0 Then SQLiteCDeclCallbackNoCaseCollating = SQLiteCreateCDeclCallback(AddressOf SQLiteFunctionNoCaseCollating, 20)
End If
SQLiteRefCount = SQLiteRefCount + 1
End Sub

Public Sub SQLiteRelease()
SQLiteRefCount = SQLiteRefCount - 1
If SQLiteRefCount = 0 Then
    stub_sqlite3_shutdown
    If SQLiteCDeclCallbackLowerUpper <> 0 Then VirtualFree ByVal SQLiteCDeclCallbackLowerUpper, 0, MEM_RELEASE: SQLiteCDeclCallbackLowerUpper = 0
    If SQLiteCDeclCallbackLike <> 0 Then VirtualFree ByVal SQLiteCDeclCallbackLike, 0, MEM_RELEASE: SQLiteCDeclCallbackLike = 0
    If SQLiteCDeclCallbackNoCaseCollating <> 0 Then VirtualFree ByVal SQLiteCDeclCallbackNoCaseCollating, 0, MEM_RELEASE: SQLiteCDeclCallbackNoCaseCollating = 0
End If
End Sub

Private Function SQLiteCreateCDeclCallback(ByVal Address As Long, ByVal cbParam As Byte) As Long
' Thanks to Paul Caton's Call CDECL source code.
Dim ASMWrapper As Long
ASMWrapper = VirtualAlloc(ByVal 0&, 28, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
Dim ASM(0 To 2) As Currency
ASM(0) = 465203369712025.6232@
ASM(1) = -140418483381718.8329@
ASM(2) = -4672484613390.9419@
CopyMemory ByVal ASMWrapper, ByVal VarPtr(ASM(0)), 24
CopyMemory ByVal UnsignedAdd(ASMWrapper, 24), &HC30672, 4
CopyMemory ByVal UnsignedAdd(ASMWrapper, 10), UnsignedSub(UnsignedSub(Address, ASMWrapper), 14), 4
CopyMemory ByVal UnsignedAdd(ASMWrapper, 16), cbParam, 1
SQLiteCreateCDeclCallback = ASMWrapper
End Function

Public Sub SQLiteOverloadBuiltinFunctions(ByVal hDB As Long)
Const STR_LOWER_UTF8 As Currency = 35335066.8108@
Const STR_UPPER_UTF8 As Currency = 35335020.9621@
Const STR_LIKE_UTF8 As Currency = 116256.1868@
Const STR_NOCASE_UTF8 As Currency = 7622387953.2366@
If hDB <> 0 Then
    If SQLiteCDeclCallbackLowerUpper <> 0 Then
        stub_sqlite3_create_function_v2 hDB, VarPtr(STR_LOWER_UTF8), 1, SQLITE_UTF8 Or SQLITE_DETERMINISTIC, 0, SQLiteCDeclCallbackLowerUpper, 0, 0, 0
        stub_sqlite3_create_function_v2 hDB, VarPtr(STR_LOWER_UTF8), 1, SQLITE_UTF16 Or SQLITE_DETERMINISTIC, 0, SQLiteCDeclCallbackLowerUpper, 0, 0, 0
        stub_sqlite3_create_function_v2 hDB, VarPtr(STR_UPPER_UTF8), 1, SQLITE_UTF8 Or SQLITE_DETERMINISTIC, 1, SQLiteCDeclCallbackLowerUpper, 0, 0, 0
        stub_sqlite3_create_function_v2 hDB, VarPtr(STR_UPPER_UTF8), 1, SQLITE_UTF16 Or SQLITE_DETERMINISTIC, 1, SQLiteCDeclCallbackLowerUpper, 0, 0, 0
    End If
    If SQLiteCDeclCallbackLike <> 0 Then stub_sqlite3_create_function_v2 hDB, VarPtr(STR_LIKE_UTF8), 2, SQLITE_UTF8 Or SQLITE_DETERMINISTIC, 0, SQLiteCDeclCallbackLike, 0, 0, 0
    If SQLiteCDeclCallbackNoCaseCollating <> 0 Then stub_sqlite3_create_collation_v2 hDB, VarPtr(STR_NOCASE_UTF8), SQLITE_UTF8, 0, SQLiteCDeclCallbackNoCaseCollating, 0
End If
End Sub

Public Function SQLiteFunctionLowerUpper(ByVal pCtx As Long, ByVal cArg As Long, ByVal pArgValue As Long) As Long
If cArg = 1 Then
    Dim pValue As Long
    CopyMemory pValue, ByVal pArgValue, 4
    Dim Text As String, cbText As Long
    cbText = stub_sqlite3_value_bytes16(pValue)
    If cbText > 0 Then
        If stub_sqlite3_user_data(pCtx) = 0 Then
            Text = LCase$(SQLiteUTF16PtrToStr(stub_sqlite3_value_text16(pValue), cbText / 2)) & vbNullChar
        Else
            Text = UCase$(SQLiteUTF16PtrToStr(stub_sqlite3_value_text16(pValue), cbText / 2)) & vbNullChar
        End If
        stub_sqlite3_result_text16 pCtx, StrPtr(Text), -1, SQLITE_TRANSIENT
    Else
        stub_sqlite3_result_text16 pCtx, 0, 0, SQLITE_STATIC
    End If
End If
End Function

Public Function SQLiteFunctionLike(ByVal pCtx As Long, ByVal cArg As Long, ByVal pArgValue As Long) As Long
If cArg = 2 Then
    Dim pValue(0 To 1) As Long
    CopyMemory pValue(0), ByVal pArgValue, 8
    If stub_sqlite3_value_bytes(pValue(0)) > SQLITE_MAX_LIKE_PATTERN_LENGTH Then
        stub_sqlite3_result_error_toobig pCtx
        Exit Function
    End If
    Dim szPattern As String, szString As String
    szPattern = SQLiteUTF8PtrToStr(stub_sqlite3_value_text(pValue(0)), stub_sqlite3_value_bytes(pValue(0)))
    szString = SQLiteUTF8PtrToStr(stub_sqlite3_value_text(pValue(1)), stub_sqlite3_value_bytes(pValue(1)))
    Dim Pos As Long, Length As Long
    Length = Len(szPattern)
    Pos = Length
    Do While Pos > 0
        Select Case Mid$(szPattern, Pos, 1)
            Case "_"
                Mid$(szPattern, Pos, 1) = "?"
            Case "%"
                Mid$(szPattern, Pos, 1) = "*"
        End Select
        Pos = Pos - 1
    Loop
    If szString Like szPattern Then ' Option Compare Text
        stub_sqlite3_result_int pCtx, 1
    Else
        stub_sqlite3_result_int pCtx, 0
    End If
End If
End Function

Public Function SQLiteFunctionNoCaseCollating(ByVal pNotUsed As Long, ByVal nKey1 As Long, ByVal pKey1 As Long, ByVal nKey2 As Long, ByVal pKey2 As Long) As Long
SQLiteFunctionNoCaseCollating = StrComp(SQLiteUTF8PtrToStr(pKey1, nKey1), SQLiteUTF8PtrToStr(pKey2, nKey2), vbTextCompare)
End Function

Public Function SQLiteBlobToByteArray(ByVal Ptr As Long, ByVal Size As Long) As Variant
If Ptr <> 0 And Size > 0 Then
    Dim B() As Byte
    ReDim B(0 To (Size - 1)) As Byte
    CopyMemory B(0), ByVal Ptr, Size
    SQLiteBlobToByteArray = B()
Else
    SQLiteBlobToByteArray = Null
End If
End Function

Public Function SQLiteUTF8PtrToStr(ByVal Ptr As Long, Optional ByVal Size As Long = -1) As String
If Ptr <> 0 Then
    If Size = -1 Then Size = lstrlenA(Ptr)
    If Size > 0 Then
        Dim Length As Long
        Length = MultiByteToWideChar(CP_UTF8, 0, Ptr, Size, 0, 0)
        If Length > 0 Then
            SQLiteUTF8PtrToStr = Space$(Length)
            MultiByteToWideChar CP_UTF8, 0, Ptr, Size, StrPtr(SQLiteUTF8PtrToStr), Length
        End If
    End If
End If
End Function

Public Function SQLiteUTF16PtrToStr(ByVal Ptr As Long, Optional ByVal Size As Long = -1) As String
If Ptr <> 0 Then
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
