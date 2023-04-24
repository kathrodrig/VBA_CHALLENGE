Sub Alpha():

For Each ws In Worksheets

    'first step, label your rows ms katherine
    ws.Range("i1") = "Ticker"
    ws.Range("j1") = "Yearly Change"      'YC
    ws.Range("k1") = "Percent Change"    'PC
    ws.Range("l1") = "Stock Volume"        'SV

    'also label the rows for functionality
    ws.Range("p1") = "Ticker"
    ws.Range("q1") = "Value"
    ws.Range("o2") = "Greates % Increase"
    ws.Range("o3") = "Greates % Decrease"
    ws.Range("o4") = "Greatest Total Volume"

    'make your variables, what variables will you use?
    Dim ticker As String
    Dim opening As Double
            opening = ws.Range("c2") 'this is gonna change later REMEMBER
    Dim closing As Double

    'what are variables are you going to figure out?
    Dim yc As Double
    Dim pc As Double
    Dim sv As LongLong 'use long long bc # is big
            sv = 0
        
    'for  the functioanality part
    Dim GIV As Double
    Dim GDV As Double
    Dim GTV As LongLong
    
    Dim tickGIV As String
    Dim tickGDV As String
    Dim tickGTV As String
    
    Summary = 2 'this will also change later
    

        'start loop
        'LOOK!! LOOK AT credit card example
        For i = 2 To 760000

         'say that if this doesnt = that, THEN  so that whenever the ticker changes, it doesn't  add to what we
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'the ticker in this cell...
                ticker = ws.Cells(i, 1).Value
                yc = ws.Cells(i, 6).Value - opening
                pc = yc / opening
                sv = sv + ws.Cells(i, 7).Value
        
                'summary just means the result
                ws.Cells(Summary, 9).Value = ticker
                ws.Cells(Summary, 10).Value = yc
                ws.Cells(Summary, 11).Value = pc
                ws.Cells(Summary, 11).NumberFormat = "#.##%"
                ws.Cells(Summary, 12).Value = sv
                
                'make sure it goes to the new row
                sv = 0
                ticker = " "
                Summary = Summary + 1
                opening = ws.Cells(i + 1, 3).Value
        
            Else
                sv = sv + ws.Cells(i, 7).Value
    
            End If
    
        'now some  formatting! gotta make it look real pretty so look at the vba color index
         'red =3  , green = 4
 
        'if the yearly change is less than 0...
        If ws.Cells(Summary, 10).Value < 0 Then
            ws.Cells(Summary, 10).Interior.ColorIndex = 3 'MAKE IT RED
                
            Else    'IF NOT...
                ws.Cells(Summary, 10).Interior.ColorIndex = 4 'MAKE IT GREEN
                
            End If 'done :)
        
        Next i
                
        'finding the greatest % increase + greatest % decrease + greatest total
      
            'kinda like the max and min functions in excel, look at vba cloud recording
            'Application.WorksheetFunction.[insert function here]range( xyz)
            GIV = Application.WorksheetFunction.Max(ws.Range("k:K"))
            GDV = Application.WorksheetFunction.Min(ws.Range("k:K"))
            GTV = Application.WorksheetFunction.Max(ws.Range("l:l"))
 
                'now, set where the results will show up
                ws.Range("q2") = GIV
                ws.Range("q3") = GDV
                ws.Range("q4") = GTV
            
                'got results! now time to do some formatting again
                'screenshot shows we need GIV & GDV to be %
                'GTV has to be in scientific notation
                'use .numberformat again
                ws.Range("q2").NumberFormat = "#.##%"
                ws.Range("q3").NumberFormat = "#.##%"
                ws.Range("q4").NumberFormat = "#.##E+##"

  
        'now the ticker
        'use excels xlookup function with application.worksheetfunction.(excelfunction)(range(xyz))
        tickGIV = Application.WorksheetFunction.XLookup(ws.Range("q2"), ws.Range("k:k"), ws.Range("i:i"), "N/A")
        tickGDV = Application.WorksheetFunction.XLookup(ws.Range("q3"), ws.Range("k:k"), ws.Range("i:i"), "N/A")
        tickGTV = Application.WorksheetFunction.XLookup(ws.Range("q4"), ws.Range("L:L"), ws.Range("i:i"), "N/A")

            'place the ticker row now
            ws.Range("p2") = tickGIV
            ws.Range("p3") = tickGDV
            ws.Range("p4") = tickGTV
 
Next ws

End Sub

'now go back and figure out how to make this show up on every sheet lol
'look at: census activity
'for each ws in work sheets, add ws before ever cells and every range, add next ws


