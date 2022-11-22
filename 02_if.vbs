num_1 = 400
num_2 = 400
num_3 = 400


If num_1 = num_2 And num_1 = num_3 And num_2 = num_3 Then
    WScript.Echo("--- EQUILATERAL TRIANGLE ---")
ElseIf num_1 = num_2 Or num_1 = num_3 Or num_2 = num_3 Then
    WScript.Echo("--- ISOSCELES TRIANGLE ---")
Else
    WScript.Echo("--- SCALENE TRIANGLE---")
End If


' If num_1 > num_2 Then
'     MsgBox num_1 & " is more than " & num_2        

' ElseIf num_1 < num_2 Then
'     WScript.Echo(num_2 & " is more than " & num_1)

' ElseIf num_1 = num_2 Then
'     WScript.Echo(num_2 & " is equal " & num_1)

' End If

    