Attribute VB_Name = "TextCompareOperators"
Option Explicit
Option Compare Text

' The Option Compare Text statement should be avoided as it makes procedures behave differently and slows down all string comparison operations.
' However, the only way to get some operators to perform in a case-insensitive manner is to use the Option Compare Text statement.
' In this case, the procedures that require this statement should be isolated in a separate module.

Public Function TextCompareLike(ByRef szString As String, ByRef szPattern As String) As Boolean
TextCompareLike = (szString Like szPattern)
End Function
