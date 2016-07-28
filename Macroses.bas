Public Function AbbrFIOSpace(surname As String, firstname As String, patronymic As String) As String
Dim res As String
res = res + CorrectCasing(surname)
res = res + " "
res = res + UCase(Left(firstname, 1)) + "."
'res = res + " "
If patronymic = "" Then
AbbrFIOSpace = res
Else
res = res + " "
res = res + UCase(Left(patronymic, 1)) + "."
AbbrFIOSpace = res
End If



End Function
Public Function AbbrFIONoSpace(surname As String, firstname As String, patronymic As String) As String
Dim res As String
res = res + CorrectCasing(surname)
res = res + " "
res = res + UCase(Left(firstname, 1)) + "."
'res = res + " "
If patronymic = "" Then
AbbrFIONoSpace = res
Else
res = res + UCase(Left(patronymic, 1)) + "."
AbbrFIONoSpace = res
End If


End Function

Public Function CorrectCasing(surname As String) As String
CorrectCasing = surname

End Function
