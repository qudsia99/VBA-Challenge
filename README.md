# VBA-Challenge

Sub stockbest()
    For Each ws In Worksheets '[Running loop in all worksheets]
        WorksheetName = ws.Name
        finalrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        '[Declaring all necessary variables]
        Dim stock_name As String
        Dim stock_tot As Double
        stock_tot = 0
        Dim row_num As Integer
        row_num = 2
        
        ' [Labeling The New Columns]
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        For i = 2 To finalrow
            '[Ticker and Total Stock Volume]
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then '[if next cell row does not match pervious cell row, then do following]
                stock_name = ws.Cells(i, 1).Value '[Assigning stock name to variable]
                ws.Range("I" & row_num).Value = stock_name
                stock_tot = stock_tot + ws.Cells(i, 7).Value
                ws.Range("L" & row_num).Value = stock_tot '[Assigning stock total to variable]
                row_num = row_num + 1
                stock_tot = 0 '[Resetting stock total so it doesn't add new values to older totals]
            Else
                stock_tot = stock_tot + ws.Cells(i, 7).Value '[If cells are the same, asking loop to add to the total]
            End If
        Next i
        
'===============
' [Yearly Change and Percentage Form]

        '[Declaring variables]
        Dim open_v As Single
        Dim closing_v As Single
        Dim yearly_c As Single
        Dim row_nu As Integer
        row_nu = 2
        
        For i = 2 To finalrow
            If ws.Cells(i, 2).Value = (ws.Name + "0102") Then  '[Finding the cells matching the opening and closing dates]
                open_v = ws.Cells(i, 3).Value
            ElseIf ws.Cells(i, 2).Value = (ws.Name + "1231") Then
                closing_v = ws.Cells(i, 6).Value
                yearly_c = closing_v - open_v  '[Applying a mathematical function to find the change value]
                ws.Cells(row_nu, 10).Value = yearly_c
                ws.Cells(row_nu, 10).NumberFormat = "#,##0.00"   '[Formatting to 2 decimal places]
                
                If open_v <> 0 Then
                    ws.Cells(row_nu, 11).Value = yearly_c / open_v
                    ws.Cells(row_nu, 11).NumberFormat = "0.00%"  '[Applying conditions for percentage]
                Else
                    ws.Cells(row_nu, 11).Value = 0
                    ws.Cells(row_nu, 11).NumberFormat = "0.00%"
                End If
                row_nu = row_nu + 1 '[Incrementing row number to go to next row after completing one loop]
            End If
        Next i
        
'================

        '[Cell formatting]
                Dim cell As Range
                For Each cell In Range("J2:J3001")
                    If cell.Value > 0 Then
                    cell.Interior.Color = RGB(0, 255, 0) ' [Green]
                     Else: cell.Interior.Color = RGB(255, 0, 0) '[Red]
                     End If
                Next cell
                
                
'==============

        '[Functionality Segment]
        
        '[Labelling the appropriate titles]
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        '[finding, MAX, MIN and GREATEST VOL. Values]
        
        Dim TV As Double '[TV = Total Volume]
        Dim maxi As Double '[maxi = Greatest percentage change]
        Dim mini As Double '[mini =greatest percentage decrease]
        Dim data_r As Range
        Set data_r = Range("L1:L3001")
        Dim data_2r As Range
        Set data_2r = Range("K1:K3001")
        
        'TV = Application.WorksheetFunction.max(data_r)
        '         TV = ws.Cells(4, 17).Value        '[Printing the Total Volume on table]
        '         ws.Cells(i, 9).Value = ws.Cells(4, 16).Value   '[Printing the corresponding tickers name]
        'maxi = Application.WorksheetFunction.max(data_2r)
        '     max = ws.Cells(2, 17).Value
        '     ws.Cells(i, 9).Value = Cells(2, 16).Value
        'mini = Application.WorksheetFunction.min(data_2r)
        '     min = ws.Cells(3, 17).Value
        '     ws.Cells(i, 9).Value = Cells(3, 16).Value
             
       
        
        
        
    Next ws
End Sub
