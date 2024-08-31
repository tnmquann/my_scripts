Attribute VB_Name = "ChamCoTrongSo"
Function MatchScoreWeighted(startCell As Range, refStartCell As Range, numQuestions As Integer) As Double
Attribute MatchScoreWeighted.VB_Description = "Calculates the weighted average score for groups of 4 cells"
Attribute MatchScoreWeighted.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim matchCount As Integer
    Dim totalScore As Double
    Dim i As Integer, j As Integer

    ' Loop through each question
    For i = 0 To numQuestions - 1
        matchCount = 0
        ' Loop through the 4 cells for each question
        For j = 0 To 3
            If startCell.Offset(0, 4 * i + j).Value = refStartCell.Offset(0, 4 * i + j).Value Then
                matchCount = matchCount + 1
            End If
        Next j
        
        ' Assign scores based on the number of matchess
        Select Case matchCount
            Case 4
                totalScore = totalScore + 1
            Case 3
                totalScore = totalScore + 0.5
            Case 2
                totalScore = totalScore + 0.25
            Case 1
                totalScore = totalScore + 0.1
            Case Else
                totalScore = totalScore + 0
        End Select
    Next i

    ' Calculate the weighted average
    MatchScoreWeighted = totalScore / numQuestions
End Function


