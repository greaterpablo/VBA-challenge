Attribute VB_Name = "Módulo1"
Sub ticker_()
      
        Dim Ticker_Name As String
            
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        
       
        Dim Open_Price As Double
        Open_Price = 0
        
        Dim Close_Price As Double
        Close_Price = 0
        
        Dim Delta_Price As Double
        Delta_Price = 0
        
        Dim Delta_Percent As Double
        Delta_Percent = 0
       
           
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

       
               
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Yearly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Stock Volume"
        
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(3, 15).Value = "Greatest % Decrease"
            Cells(4, 15).Value = "Greatest Total Volume"
            Cells(1, 16).Value = "Ticker"
            Cells(1, 17).Value = "Value"
       
             
        
        Open_Price = Cells(2, 3).Value
        
        
        For i = 2 To Lastrow
        
      
           
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
               
                Ticker_Name = Cells(i, 1).Value
                
                
                Close_Price = Cells(i, 6).Value
                Delta_Price = Close_Price - Open_Price
             
                If Open_Price <> 0 Then
                    Delta_Percent = (Delta_Price / Open_Price) * 100
                
                End If
                
               
                Total_Ticker_Volume = Total_Ticker_Volume + Cells(i, 7).Value
              
                
             
                Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                Range("J" & Summary_Table_Row).Value = Delta_Price
                
                If (Delta_Price > 0) Then
                    
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Delta_Price <= 0) Then
                    
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 
                Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")
               
                Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
              
                Summary_Table_Row = Summary_Table_Row + 1
                
                Delta_Price = 0
                
                Close_Price = 0
                
                Open_Price = Cells(i + 1, 3).Value
              
                               
                             
              
                Delta_Percent = 0
                Total_Ticker_Volume = 0
                
            
            
         
            Else
                
                Total_Ticker_Volume = Total_Ticker_Volume + Cells(i, 7).Value
            End If
            
      
        Next i

End Sub
