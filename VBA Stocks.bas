Attribute VB_Name = "Module1"
        Sub Markets()


        Dim i As Double
        Dim k As Double
        Dim Table_Row As Double
        Dim ws_num As Integer
        Dim starting_ws As Worksheet
        Dim ticker As String
        Dim open_price As Long
        Dim close_price As Long
        Dim totalstock As Double
        Dim sheet As Integer
        Dim n As String
        Dim inctick As String
        Dim mintick As String
        Dim voltick As String
        Dim increase As Double
        Dim min As Double
        Dim volume As Double



        Set starting_ws = ActiveSheet

        ws_num = ThisWorkbook.Worksheets.Count

        For sheet = 1 To ws_num

        ThisWorkbook.Worksheets(sheet).Activate
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "% Change"
        Range("L1").Value = "Total Stock Volume"
        totalstock = 0
        close_price = 0
        open_price = 0
        Table_Row = 2
        jan1 = 20160101
        dec30 = 20161230
        chng = 0
        perchng = 0
        Range("n2") = "Greatest % Increase"
        Range("N3") = "Greatest % Decrease"
        Range("N4") = "Greatest Total Volume"
        Range("O1") = "Ticker"
        Range("P1") = "Value"



        last_row = Cells(Rows.Count, 1).End(xlUp).Row

            For i = 2 To last_row
            

                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    
                    ticker = Cells(i, 1).Value
                    totalstock = Cells(i, 7).Value + totalstock
                    close_price = Cells(i, 6).Value
                    chng = (close_price - open_price)
                    
                    If chng = 0 Or open_price = 0 Then
                    
                    perchng = 0
                    Else
                    perchng = (chng / open_price)
                    End If
                                
                    
                    Range("I" & Table_Row).Value = ticker
                    Range("J" & Table_Row).Value = chng
                    Range("K" & Table_Row).Value = perchng
                    Range("k" & Table_Row).NumberFormat = "0%"
                    Range("L" & Table_Row).Value = totalstock

                    If chng > 0 Then
                    
                    Range("j" & Table_Row).Interior.ColorIndex = 4
                    
                    ElseIf chng < 0 Then
                    
                    Range("j" & Table_Row).Interior.ColorIndex = 3
                    
                    End If
                
                    
                    totalstock = 0
                    Table_Row = Table_Row + 1
                    open_price = 0
                    close_price = 0
                    
                Else
                
                    totalstock = totalstock + Cells(i, 8).Value
                    If Cells(i, 2).Value = jan1 Then
                        open_price = Cells(i, 3).Value
                    End If

                End If

            Next i

    
        For k = 2 To 300


                 Range("P2").NumberFormat = "0%"
                 Range("P3").NumberFormat = "0%"
            
                
                
                If Cells(k, 11).Value > increase Then
                
                increase = Cells(k, 11).Value
                inctick = Cells(k, 9).Value
      
                
                End If

                If Cells(k, 11).Value < min Then
                min = Cells(k, 11).Value
                mintick = Cells(k, 9).Value
                
                
                
                End If
            
                If Cells(k, 12).Value > volume Then
                volume = Cells(k, 12).Value
                volumetick = Cells(k, 9).Value

                End If
                

     Next k


                Range("O2").Value = inctick
                Range("O3").Value = mintick
                Range("O4").Value = volumetick
                Range("P2").Value = increase
                Range("P3").Value = min
                Range("P4").Value = volume


        Next sheet

        starting_ws.Activate

        End Sub


