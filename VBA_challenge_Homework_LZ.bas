Attribute VB_Name = "Module1"
Sub StockVBA():

For Each ws In Worksheets

Dim Ticker As String
Dim Volume_Counter As Double
Dim Summary_Table_Row As Integer
Dim Open_Year As Double
Dim Close_Year As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Summary_Table_Row = 2
Volume_Counter = 0
Open_Year = ws.Cells(2, 3).Value
    
    For i = 2 To lastrow
                                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                    Ticker = ws.Cells(i, 1).Value
                                    Volume_Counter = Volume_Counter + ws.Cells(i, 7).Value
                                    Close_Year = ws.Cells(i, 6).Value
                                    Yearly_Change = Close_Year - Open_Year
                                    
                                                        If Open_Year <> 0 Then
                                                            Percent_Change = (Yearly_Change / Open_Year) * 100
                                                        Else
                                                            MsgBox ("Fix <open> field manually and save the spreadsheet.")
                                                        End If
                                                        
                                    ws.Range("I" & Summary_Table_Row).Value = Ticker
                                    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                                    ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
                                    ws.Range("L" & Summary_Table_Row).Value = Volume_Counter
                                                'If ws.Cells(i, 11).Value = Min
                                                'ws.Range("P" & Summary_Table_Row).Value = Min(Percent_Change)
                                                'ws.Range("Q" & Summary_Table_Row).Value = Max(Volume_Counter)
                                    Summary_Table_Row = Summary_Table_Row + 1
                                    Volume_Counter = 0
                                    Open_Year = ws.Cells(i + 1, 3).Value
                        
                        Else
                            Volume_Counter = Volume_Counter + ws.Cells(i, 7).Value
                                
                        End If
                                             
                                    If ws.Cells(i, 10).Value < 0 Then
                                    ws.Cells(i, 10).Interior.ColorIndex = 3
                                    
                                    ElseIf ws.Cells(i, 10).Value >= 0 Then
                                    ws.Cells(i, 10).Interior.ColorIndex = 4
                                    
                                    Else
                                    ws.Cells(i, 10).Interior.ColorIndex = xlNone
                                    
                                    End If
            
                    
            Next i
            
    
    Next ws
    
End Sub

