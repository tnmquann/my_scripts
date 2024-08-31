Attribute VB_Name = "Module1"
Function UpdateConstantCellRange(ByVal initialCell As Range, ByVal increment As Integer) As Range
    Dim rowNumber As Integer
    Dim columnNumber As Integer
    
    ' Get the row and column numbers from the initial cell
    rowNumber = initialCell.Row
    columnNumber = initialCell.Column
    
    ' Calculate the ending column based on the increment value
    Dim endingColumn As Integer
    endingColumn = columnNumber + increment - 1
    
    ' Convert column numbers to column letters
    Dim startingColumnLetter As String
    Dim endingColumnLetter As String
    startingColumnLetter = Split(Cells(1, columnNumber).Address, "$")(1)
    endingColumnLetter = Split(Cells(1, endingColumn).Address, "$")(1)
    
    ' Build the range address string
    Dim rangeAddress As String
    rangeAddress = "$" & startingColumnLetter & "$" & rowNumber & ":$" & endingColumnLetter & "$" & rowNumber
    
    ' Return the range of cells
    Set UpdateConstantCellRange = Range(rangeAddress)
End Function
