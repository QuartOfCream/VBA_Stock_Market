Attribute VB_Name = "Module1"
Sub VBA_Alpha_Confusion()

    Dim Ticker As String
    Dim Ticker_Volume As Double
    Dim Final_Column As Long
    Dim FinalRow As Long
    Dim i As Long
    Dim j As Integer
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Price_Change As Double
    Dim Price_Change_Percent As Double
    

    For Each ws In Worksheets

'Set Names of Headers
    ws.Range("H1").Value = "Ticker"
    ws.Range("I1").Value = "Yearly Change"
    ws.Range("J1").Value = "Percent Change"
    ws.Range("K1").Value = "Total Stock Volume"
    

   
'Make sure everything is starting out at zero
    Open_Price = 0
    Close_Price = 0
    Price_Change = 0
    Price_Change_Percent = 0
    Ticker_Volume = 0
    j = 0
    
     
'Start Final_Column at 2 to avoid the header
    Final_Column = 2
    
    
    For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
'Figure out where the final row is
   ' FinalRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
'If Ticker is not equal to the next unique ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Final_Column = Final_Column + 1
        
         Price_Change = Close_Price - Open_Price
         
    'Define the Ticker variable
        Ticker = ws.Cells(i, 1).Value
        
        'Figure out what the Close Price parameter and Percent Change
            Close_Price = ws.Cells(i, 6).Value
        
        'Print the info I tried two ways!
                 '****************
        ws.Range("H" & Final_Column).Value = Ticker
        ws.Cells(Final_Column, 10).Value = Price_Change
        
        
        'Price Change with Formatting
        ws.Range("I" & Final_Column).Value = Price_Change
        
        If Price_Change > 0 Then
        
            ws.Range("I" & Final_Column).Interior.ColorIndex = 4
        Else
        
            ws.Range("I" & Final_Column).Interior.ColorIndex = 3
        End If
        
        
        'Percent Change code
        If Open_Price <> Close_Price Then
            Price_Change_Percent = (Price_Change / Close_Price)
        Else
            Price_Change_Percent = 0
        End If
        ws.Range("J" & Final_Column).Value = Price_Change_Percent
        ws.Range("J" & Final_Column).NumberFormat = "0.00%"


        'Add total ticker volume
        Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
        ws.Range("K" & Final_Column).Value = Ticker_Volume
        
    'Tell Excel where to put the values for Ticker
        'ws.Cells(Final_Column, 9).Value = Ticker
        
    
       
        
    'If the Open Price equals Zero
            
         
     End If
     
     
        
    
        
        
     
     Next i
     
     Next ws
     
     




End Sub


