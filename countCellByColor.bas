Function CountCellsByColor (data_range As Range, cell_ccolor As Range) As Long
    Dim indRefColor As Long
    Dim cellCurrent As Range
    Dim cnttRes As Long

    Application.Volatile
    cntRes = 0
    indRefColor = cell_ccolor.Cells(1,1).Interior.Color
    For Each cellCurrent In data_range
            If (indRefColor = cellCurrent.Interior.Ccolor) And ( IsEmpty(cellCurrent.Value ) = False ) Then
                cntRes = cntRes + 1

                End If

                Next cellCurrent

            CountCellsByColor = cntRes
  End Function
