Attribute VB_Name = "Module1"
Sub StockAnalysis():
 
    Dim StockAnalysis As Worksheet
    Dim Stock_Table As Boolean
    Dim Spreadsheet As Boolean
    
    Stock_Table = False
    Spreadsheet = True
   
    For Each StockAnalysis In Worksheets
     
        Dim Ticker_Name As String
        Ticker_Name = " "
        
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
    
        Dim MAX_TICKER_NAME As String
        MAX_TICKER_NAME = " "
        Dim MIN_TICKER_NAME As String
        MIN_TICKER_NAME = " "
        Dim MAX_PERCENT As Double
        MAX_PERCENT = 0
        Dim MIN_PERCENT As Double
        MIN_PERCENT = 0
        Dim MAX_VOLUME_TICKER As String
        MAX_VOLUME_TICKER = " "
        Dim MAX_VOLUME As Double
        MAX_VOLUME = 0

        Dim Summary_Table As Long
        Summary_Table = 2
         
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = StockAnalysis.Cells(Rows.Count, 1).End(xlUp).Row
    
        If Stock_Table Then
            
            StockAnalysis.Range("I1").Value = "Ticker"
            StockAnalysis.Range("J1").Value = "Yearly Change"
            StockAnalysis.Range("K1").Value = "Percent Change"
            StockAnalysis.Range("L1").Value = "Total Stock Volume"
            
            StockAnalysis.Range("O2").Value = "Greatest % Increase"
            StockAnalysis.Range("O3").Value = "Greatest % Decrease"
            StockAnalysis.Range("O4").Value = "Greatest Total Volume"
            StockAnalysis.Range("P1").Value = "Ticker"
            StockAnalysis.Range("Q1").Value = "Value"
        Else
            
            Stock_Table = True
        End If
            
        Open_Price = StockAnalysis.Cells(2, 3).Value
        
        For i = 2 To Lastrow
    
            If StockAnalysis.Cells(i + 1, 1).Value <> StockAnalysis.Cells(i, 1).Value Then
            
               
                Ticker_Name = StockAnalysis.Cells(i, 1).Value
                Close_Price = StockAnalysis.Cells(i, 6).Value
                Delta_Price = Close_Price - Open_Price
                
                If Open_Price <> 0 Then
                    Delta_Percent = (Delta_Price / Open_Price) * 100
                Else
                    
                    Delta_Percent = 0
                
                End If
                  
                Total_Ticker_Volume = Total_Ticker_Volume + StockAnalysis.Cells(i, 7).Value
        
               
                StockAnalysis.Range("I" & Summary_Table).Value = Ticker_Name
                
                StockAnalysis.Range("J" & Summary_Table).Value = Delta_Price
                
                If (Delta_Price > 0) Then
                   
                    StockAnalysis.Range("J" & Summary_Table).Interior.ColorIndex = 4
                ElseIf (Delta_Price <= 0) Then
                 
                    StockAnalysis.Range("J" & Summary_Table).Interior.ColorIndex = 3
                End If
                
                 
                StockAnalysis.Range("K" & Summary_Table).Value = (CStr(Delta_Percent) & "%")
                
                StockAnalysis.Range("L" & Summary_Table).Value = Total_Ticker_Volume
                
                
                Summary_Table = Summary_Table + 1
                
                Delta_Price = 0
                
                Close_Price = 0
              
                Open_Price = StockAnalysis.Cells(i + 1, 3).Value
                
              
                If (Delta_Percent > MAX_PERCENT) Then
                    MAX_PERCENT = Delta_Percent
                    MAX_TICKER_NAME = Ticker_Name
                ElseIf (Delta_Percent < MIN_PERCENT) Then
                    MIN_PERCENT = Delta_Percent
                    MIN_TICKER_NAME = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > MAX_VOLUME) Then
                    MAX_VOLUME = Total_Ticker_Volume
                    MAX_VOLUME_TICKER = Ticker_Name
                End If
                
                Delta_Percent = 0
                Total_Ticker_Volume = 0
                
    
      
            Else
                
                Total_Ticker_Volume = Total_Ticker_Volume + StockAnalysis.Cells(i, 7).Value
            End If
            
    
        Next i

      
            If Not Spreadsheet Then
            
                StockAnalysis.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
                StockAnalysis.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
                StockAnalysis.Range("P2").Value = MAX_TICKER_NAME
                StockAnalysis.Range("P3").Value = MIN_TICKER_NAME
                StockAnalysis.Range("Q4").Value = MAX_VOLUME
                StockAnalysis.Range("P4").Value = MAX_VOLUME_TICKER
                
            Else
                Spreadsheet = False
            End If
            
            Next StockAnalysis

 End Sub

