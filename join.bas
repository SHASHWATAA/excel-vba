Attribute VB_Name = "Module1"
Function Join(range, Optional delimiter = "")
Dim temp


For Each c In range:
If Not (c = "") Then
    temp = temp + c + delimiter
End If
Next c

Join = Left(temp, Len(temp) - Len(delimiter))


End Function

Function pythonJoin(range)

delimiter = "', '"
temp = "['"

For Each c In range:
If Not (c = "") Then
    temp = temp + c + delimiter
End If
Next c

pythonJoin = Left(temp, Len(temp) - Len(delimiter)) + "']"


End Function


