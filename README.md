# Learning VB (Visual Basic)

Just one of the things I'm learning. https://github.com/hchiam/learning

TODO: find my high school VB/VBA projects (a game, a neural network simulation, and a gravity/atoms interaction simulation).

## Counting cells by background color:

To get something like this in an Excel formula: (first range is cells to count, second range is cell to use color)

```xlsm
=GetColorCount(B18:G19,H20)
```

you need this: https://trumpexcel.com/count-colored-cells-in-excel/

```vb
'Code created by Sumit Bansal from https://trumpexcel.com
Function GetColorCount(CountRange As Range, CountColor As Range)
  Dim CountColorValue As Integer
  Dim TotalCount As Integer
  CountColorValue = CountColor.Interior.ColorIndex
  Set rCell = CountRange
  For Each rCell In CountRange
    If rCell.Interior.ColorIndex = CountColorValue Then
      TotalCount = TotalCount + 1
    End If
  Next rCell
  GetColorCount = TotalCount
End Function
```
