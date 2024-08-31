Attribute VB_Name = "ChamKhongTrongSo"
Function MatchScoreUnweighted(startCell As Range, refStartCell As Range, increment As Integer) As Double
Attribute MatchScoreUnweighted.VB_Description = "Calculates the unweighted match score between two cell ranges"
Attribute MatchScoreUnweighted.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim matchCount As Integer
    Dim i As Integer
    Dim totalCells As Integer
    
    ' Loop through the range of cells to compare
    For i = 0 To increment - 1
        If startCell.Offset(0, i).Value = refStartCell.Offset(0, i).Value Then
            matchCount = matchCount + 1
        End If
    Next i
    
    ' Count the number of non-empty cells in the reference range
    totalCells = Application.WorksheetFunction.CountA(refStartCell.Resize(1, increment))
    
    ' Calculate the match score
    If totalCells > 0 Then
        MatchScoreUnweighted = matchCount / totalCells
    Else
        MatchScoreUnweighted = 0
    End If
End Function


