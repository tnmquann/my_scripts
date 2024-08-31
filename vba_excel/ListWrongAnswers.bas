Attribute VB_Name = "LietKeCauSai"
Function ListWrongAnswers(startCell As Range, refStartCell As Range, wrongAnswerCell As Range, increment As Integer, compareValue As Variant) As String
    Dim i As Integer
    Dim wrongAnswers() As String
    Dim wrongAnswersCount As Integer
    
    ' Initialize variables
    wrongAnswersCount = 0
    
    ' Redimension array to hold potential wrong answers
    ReDim wrongAnswers(0 To increment - 1)
    
    ' Loop through each cell in the ranges
    For i = 0 To increment - 1
        If startCell.Offset(0, i).Value <> refStartCell.Offset(0, i).Value Then
            ' Store wrong answers in the array
            If wrongAnswerCell.Offset(0, i).Value <> compareValue Then
                wrongAnswers(wrongAnswersCount) = wrongAnswerCell.Offset(0, i).Value
                wrongAnswersCount = wrongAnswersCount + 1
            End If
        End If
    Next i
    
    ' Resize the array to fit the number of wrong answers
    ReDim Preserve wrongAnswers(0 To wrongAnswersCount - 1)
    
    ' Join the array into a single string
    ListWrongAnswers = Join(wrongAnswers, ", ")
End Function

