Private Sub Main()
Dim lol as Integer: lol = InputBox("")
Open "test.html" For Output as #1
For lol = 0 to lol - 1
Print #1, "<input type=text name=" & lol + 1 & ">"
Next
Close #1
End Sub