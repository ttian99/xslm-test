Attribute VB_Name = "SumTextValue"
Function all(Rng As Range) As String
    If Len(Rng.Value) <> 0 Then
        With CreateObject("vbscript.regexp")
            .Pattern = "[\d\+\-\*\(\)/.]+"
            .Global = True
            .MultiLine = False
            If .Test(Rng) Then
                .Pattern = "(¡¾.*?¡¿)*"
                .Global = True
                .MultiLine = False
                If .Test(Rng) Then
                    Dim str
                    all = Application.Evaluate("=" & .Replace(Rng, ""))
                Else
                    all = ""
                End If
            End If
        End With
    Else
        all = ""
    End If
End Function


Sub all_Click()
    Dim str
    str = all(Range("a3"))
    Debug.Print str
    Sheet1.Cells(10, 10) = str
End Sub

    
