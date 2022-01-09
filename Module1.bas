Attribute VB_Name = "Module1"
Sub abTest():


' Declare Current as a worksheet object variable.
         Dim Current As Worksheet
         Dim next_line As Integer
         Dim end_row As Double
         Dim stock_name As String
         Dim stock_total As Double
         
         
         
         
         
         
         ' Loop through all of the worksheets in the active workbook.
         For Each Current In Worksheets
         
         stock_total = 0
         next_line = 2
         end_row = Current.Cells.SpecialCells(xlCellTypeLastCell).Row

            ' Insert yo ur code here.
            
            For i = 2 To end_row
                If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1) Then
                
                    stock_total = stock_total + Current.Cells(i, 7).Value
                    stock_name = Current.Cells(i, 1).Value
                    Current.Cells(i, 8).Value = "Change"
                    Current.Cells(next_line, 9).Value = stock_name
                    Current.Cells(next_line, 10).Value = stock_total
                    
                    next_line = next_line + 1
                    
                    stock_total = 0
                
                Else
                
                    stock_total = stock_total + Current.Cells(i, 7).Value
                    
                    
                    
                End If
                
                
               'Range("A,i").Value = 100
                
            Next i
            
            
            
            
            
            
            'Cells(i, 2).Value = 100
            'Range("i2").Value = 100
            
            
            ' This line displays the worksheet name in a message box.
            MsgBox Current.Name
         Next






End Sub

