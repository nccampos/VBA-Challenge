Attribute VB_Name = "Module1"
Sub Stock_Analysis():

    ' Set Variable for Ticker
    Dim total As Double
    Dim i As Long
    Dim change As Single
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Single
    Dim days As Integer
    Dim dailyChange As Single
    Dim averageChange As Single

    ' Title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    ' Value
    j = 0
    total = 0
    change = 0
    start = 2

    ' Remaining Row
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To rowCount

        ' Print results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Variables and values
            total = total + Cells(i, 7).Value

            ' If Value is set to zero
            If total = 0 Then
                ' Results printed
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0

            Else
                ' Value at zero
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If

                ' Change calculation
                change = (Cells(i, 6) - Cells(start, 3))
                percentChange = change / Cells(start, 3)

                ' Next Ticker
                start = i + 1

                ' Results Printed
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).Value = percentChange
                Range("K" & 2 + j).NumberFormat = "0.00%"
                Range("L" & 2 + j).Value = total

                ' Colors set as Green for positive and Red for Negative
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select

            End If

            ' New ticker rotation
            total = 0
            change = 0
            j = j + 1
            days = 0

        ' Ticker addtional
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

End Sub


Sub VBA_Homework_Macro()
Attribute VBA_Homework_Macro.VB_ProcData.VB_Invoke_Func = " \n14"
'
' VBA_Homework_Macro Macro
'

'
    ActiveWorkbook.Save
End Sub
