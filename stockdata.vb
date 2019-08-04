Sub Worksheet()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call Stock_year
    Next ws
    Application.ScreenUpdating = True
End Sub

Sub Stock_year()

Dim Ticker As String
Dim Total_Stock_Volume As Double
Dim Summary_Table_Index As Double


Range("I1").Value = "Ticker"
Range("J1").Value = "Total_Stock_Volume"
Summary_Table_Index = 2

For i = 2 To 797711


    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

        Range("I" & 1 + Summary_Table_Index).Value = Ticker
        Range("J" & 1 + Summary_Table_Index).Value = Total_Stock_Volume
                
        Summary_Table_Index = Summary_Table_Index + 1

        Total_Stock_Volume = 0

    Else
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    End If


        
Next i

End Sub
