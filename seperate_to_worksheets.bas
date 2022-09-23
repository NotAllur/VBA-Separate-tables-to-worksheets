'Made with the help of @Cyril (https://stackoverflow.com/users/3233363/cyril), by NotALlur (https://github.com/NotAllur/VBA-Separate-tables-to-worksheets)
Sub Seperate_to_worksheets()
    With Sheets(1)
        Dim lastRow As Long:  lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        Dim i As Long
        Dim boxTitle
            boxTitle = "Seperate to worksheets"
        Dim firstFindValue As String:  firstFindValue = InputBox("Enter the starting value.", boxTitle, "Cell value")
        Dim secondFindValue As String:  secondFindValue = InputBox("Enter the ending value.", boxTitle, "Cell value")
        For i = 1 To lastRow
            Dim firstFoundCell As Range:  Set firstFoundCell = .Range(.Cells(i, 1), .Cells(lastRow, 1)).Find(what:=firstFindValue, LookIn:=xlValues, lookat:=xlPart, MatchCase:=False)
            If firstFoundCell Is Nothing Then
                Exit For
            Else
                Dim secondFoundCell As Range:  Set secondFoundCell = .Range(.Cells(firstFoundCell.Row + 1, 1), .Cells(lastRow, 1)).Find(what:=secondFindValue, LookIn:=xlValues, lookat:=xlPart, MatchCase:=False)
                If secondFoundCell Is Nothing Then Exit For
                Dim destinationSheet As Worksheet:  Set destinationSheet = ThisWorkbook.Sheets.Add
                .Range(.Rows(firstFoundCell.Row), .Rows(secondFoundCell.Row)).Copy destinationSheet.Cells(1, 1)
                i = secondFoundCell.Row - 1
                Set firstFoundCell = Nothing
                Set secondFoundCell = Nothing
            End If
        Next i
    End With
End Sub
