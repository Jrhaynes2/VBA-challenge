Sub Stock_Analysis()
    ' Define Variables
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Percent_Change As Double
    Dim MaxInc As Double
    Dim MaxDec As Double
    Dim MaxTotal As Double
    Dim MaxIncTic As Variant
    Dim MaxDecTic As Variant
    Dim MaxTotTic As Variant
    Dim LookupValue As Variant
    Dim LookupRange As Range
    Dim ws As Worksheet
    Dim Total_Stock_Volume As Double
    Dim Summary_Table_Row As Integer
    Dim i As Long
    
    For Each ws In ThisWorkbook.Worksheets
        ' Set initial values for maximum values
        MaxInc = 0
        MaxDec = 0
        MaxTotal = 0
        Total_Stock_Volume = 0
        Summary_Table_Row = 2
        
        ' Define Column Name
        With ws
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Yearly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
            .Cells(1, 16).Value = "Ticker"
            .Cells(1, 17).Value = "Value"
            .Cells(2, 15).Value = "Greatest % Increase"
            .Cells(3, 15).Value = "Greatest % Decrease"
            .Cells(4, 15).Value = "Greatest Total Volume"
            
            LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            For i = 2 To LastRow
                ' Locate change in Ticker and capture close price and format table
                If .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then
                    Ticker = .Cells(i, 1).Value
                    Total_Stock_Volume = Total_Stock_Volume + .Cells(i, 7).Value
                    Close_Price = .Cells(i, 6).Value
                    ' Add unique stock tickers
                    .Range("I" & Summary_Table_Row).Value = Ticker
                    ' Calculate yearly change and format
                    .Range("J" & Summary_Table_Row).Value = Close_Price - Open_Price
                    ' Calculate percent change and format
                    If Open_Price <> 0 Then
                        .Range("K" & Summary_Table_Row).Value = (Close_Price - Open_Price) / Open_Price
                        .Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    Else
                        .Range("K" & Summary_Table_Row).Value = "N/A"
                    End If
                    ' Sum total stock volume
                    .Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                    Summary_Table_Row = Summary_Table_Row + 1
                    Total_Stock_Volume = 0
                Else
                    ' Calculate the open price and total stock volume
                    Total_Stock_Volume = Total_Stock_Volume + .Cells(i, 7).Value
                    If .Cells(i - 1, 1).Value <> .Cells(i, 1).Value Then
                        Open_Price = .Cells(i, 3).Value
                    End If
                End If
            Next i
            
            For i = 2 To LastRow
                ' Set positive values to green
                If .Cells(i, 10).Value > 0 Then
                    .Cells(i, 10).Interior.ColorIndex = 4
                ' Set empty cells to no fill
                ElseIf .Cells(i, 10).Value = "" Then
                    .Cells(i, 10).Interior.ColorIndex = -4142
                ' Set negative values to red
                Else
                    .Cells(i, 10).Interior.ColorIndex = 3
                End If
            Next i

            ' Determine Greatest Increase %
            MaxInc = WorksheetFunction.Max(.Range("K2:K" & LastRow))
            .Range("Q2").Value = MaxInc
            .Range("Q2").NumberFormat = "0.00%"
            LookupValue = MaxInc
            Set LookupRange = .Range(("K2:K" & LastRow))
            MaxIncTic = WorksheetFunction.Match(LookupValue, LookupRange, 0)
            If Not IsError(MaxIncTic) Then
                .Range("P2").Value = .Cells(Int(MaxIncTic) + 1, 9).Value
            Else
                .Range("P2").Value = "N/A"
            End If

            ' Determine Greatest Decrease %
            MaxDec = WorksheetFunction.Min(.Range("K2:K" & LastRow))
            .Range("Q3").Value = MaxDec
            .Range("Q3").NumberFormat = "0.00%"
            LookupValue = MaxDec
            Set LookupRange = .Range(("K2:K" & LastRow))
            MaxDecTic = WorksheetFunction.Match(LookupValue, LookupRange, 0)
            If Not IsError(MaxDecTic) Then
                .Range("P3").Value = .Cells(Int(MaxDecTic) + 1, 9).Value
            Else
                .Range("P3").Value = "N/A"
            End If
            
            ' Determine Greatest Total Volume
            MaxTotal = WorksheetFunction.Max(.Range("L2:L" & LastRow))
            .Range("Q4").Value = MaxTotal
            LookupValue = MaxTotal
            Set LookupRange = .Range(("L2:L" & LastRow))
            MaxTotTic = Application.Match(LookupValue, LookupRange, 0)
            If Not IsError(MaxTotTic) Then
                .Range("P4").Value = .Cells(Int(MaxTotTic) + 1, 9).Value
            Else
                .Range("P4").Value = "N/A"
            End If

        End With
    Next ws
End Sub
