Const aaa As String = "1+324+2*3+2����������*121sdjf+ 1/232���"

Function Test(Rng As Range) As String
    With CreateObject("vbscript.regexp")
        .Pattern = "[\d\+\-\*\(\)/.]+"
        .Global = True
        .MultiLine = False
        If .Test(aaa) Then
            MsgBox Rng + "&", buttonType, "title"

        End If
    End With
End Function

