Sub Stock()

'define variable

    Dim Ticker As String
    Dim yearlyChange, percentChange As Double
    Dim totalVol As Double
    Dim ws As Worksheet
    yearlyChange = 0
    percentChange = 0
    totalVol = 0
    SummaryRow = 2
    StartRow = 2
   
'use foe ,per worksheet for per year 2018,2019,2020

For Each ws In Worksheets

' find lastrow in column A

 LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 ' creat header for sumery table
   
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percentage Change"
 ws.Cells(1, 12).Value = "Total Stock Volume"
 
 
 'defult value for totalVolume
 
 totalVol = ws.Cells(2, 7).Value
  
  'create column I and L for find TICKER and Total Volum (The total stock volume of the stock.)
    For i = 2 To LastRow

           
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
          
           totalVol = totalVol + ws.Cells(i + 1, 7).Value

            ws.Range("I" & SummaryRow).Value = Ticker
            ws.Range("L" & SummaryRow).Value = totalVol
         
            SummaryRow = SummaryRow + 1
            
             totalVol = 0
     
        Else
             totalVol = totalVol + ws.Cells(i + 1, 7).Value
         
             

        End If
        

    Next i
    
    'show message for next step
    MsgBox ("next step")
    
    'find lastrow in column I (summery table
    
     lr = ws.Cells(Rows.Count, 9).End(xlUp).Row
     
     'create and calculate yearlychange (Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.)
     'yearlychange style is grren for positive and red for negative
     'percentChange (Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.)
     
     
     For j = 2 To lr
    
       Ticker = ws.Cells(j, 9).Value
     
       StartRow = Range("A:A").Find(what:=Ticker, after:=Range("A1")).Row
       EndRow = Range("A:A").Find(what:=Ticker, after:=Range("A1"), searchdirection:=xlPrevious).Row
 
      
       
            If ws.Cells(StartRow, 1).Value <> Ticker Then
    
    
            yearlyChange = (ws.Cells(EndRow, 6).Value - ws.Cells(StartRow, 3).Value)
            
                If yearlyChange >= 0 Then
        
                ws.Cells(j, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
     
            If yearlyChange <> 0 Then
            
            percentChange = (yearlyChange / ws.Cells(StartRow, 3).Value)
            
            yearlyChange = 0
            percentChange = 0
          
            
            End If
            
           
            ws.Range("J" & j).Value = yearlyChange
            ws.Range("K" & j).Value = percentChange
            ws.Range("K" & j).Style = "Percent"
                    
          
            Else
            
            yearlyChange = (ws.Cells(EndRow, 6).Value - ws.Cells(StartRow, 3).Value)
            
              If yearlyChange >= 0 Then
        
                ws.Cells(j, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
                
            
              If yearlyChange <> 0 Then
            
            percentChange = (yearlyChange / ws.Cells(StartRow, 3).Value)
            
           
            
            End If
            
             If yearlyChange <> 0 Then
            
            percentChange = ((yearlyChange / ws.Cells(StartRow, 3).Value))
           
            
                            
            End If
          
                                   
          
            ws.Range("J" & j).Value = yearlyChange
            ws.Range("K" & j).Value = percentChange
            ws.Range("K" & j).Style = "Percent"
            
            
            End If
            
            
    Next j
    
  'show message box to continue next worksheet
  MsgBox ("calculate next year")
  
    
Next ws
  'show message box to completed all  worksheet
  
    MsgBox ("click to see result!")
  

End Sub

