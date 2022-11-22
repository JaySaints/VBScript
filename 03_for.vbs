
Dim num_min_1, num_min_2

num_1 = 420
num_2 = 2
num_3 = 303


' Take the two smollest values
For i=1 To 2
    If num_1 < num_2 And num_1 < num_3 Then
        num_min_1 = num_1

        If num_2 < num_3 Then
            num_min_2 = num_2
        Else
            num_min_2 = num_3
        End If

    ElseIf num_2 < num_3 And num_2 < num_3 Then
        num_min_1 = num_2

        If num_1 < num_3 Then
            num_min_2 = num_1
        Else
            num_min_2 = num_3
        End If

    ElseIf num_3 < num_1 And num_3 < num_1 Then
        num_min_1 = num_3

        If num_1 < num_2 Then
            num_min_2 = num_1
        Else
            num_min_2 = num_2
        End If      

    End If    
Next 'i

WScript.Echo("Two Smallest Values -> "& num_min_1 &" - "& num_min_2)




