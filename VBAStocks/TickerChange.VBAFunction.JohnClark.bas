Attribute VB_Name = "Module1"
Sub TickerChange():

        'Set Variables
        Dim Ticker As String 'Ticker Symbol
        Dim OValue As Double 'Opening Value
        Dim CValue As Double 'Closing Value
        Dim YChange As Double 'Yearly Change
        Dim PerChange As Double 'Percent Change
        Dim SValue As Double 'Stock Volume
        Dim SValueSum As Double 'Stock Volume Sum Total
    
        'Min/Max Variables
        Dim PerChangeMax As Double
        Dim PerChangeMin As Double
        Dim SValueSumMax As Double
    
    
        'Set Loop Variables
        Dim i As Long
        Dim j As Integer
        Dim LastRow As Long
        Dim k As Long
        Dim LastRow2 As Long
        Dim l As Long
        Dim m As Integer
    
    For Each ws In Worksheets
    
        'Set initial values
        j = 2
        OValue = ws.Cells(2, 3)
            'MsgBox (OValue) Good
                       
        Ticker = ws.Cells(2, 1)
            'MsgBox (Ticker) Good
            
            
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow
    
            'Get Changing Values
            CValue = ws.Cells(i, 6)
                'MsgBox (CValue) Good
            SValue = ws.Cells(i, 7)
        
            'CALCULATIONS
            YChange = (CValue - OValue)
                'Good
            If OValue = 0 Then
                PerChange = 0
            Else: PerChange = (YChange / OValue)
            End If
                
                'MsgBox (PChange) Good
            SValueSum = (SValueSum + SValue)
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Reset Variables
                OValue = ws.Cells(i + 1, 3)
                Ticker = ws.Cells(i + 1, 1)
                SValueSum = 0
                      
                j = (j + 1)
                    'MsgBox (j) good, proving we can go down a row in table
                
            Else
                'Print Values
                'j = 2
                ws.Cells(j, 9).Value = Ticker
                ws.Cells(j, 10).Value = YChange
                ws.Cells(j, 11).Value = PerChange
                ws.Cells(j, 12).Value = SValueSum
                'good
                
            End If
            
        Next i
        '-----------------------------------------------------------------
        'Min/Max Variables
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For m = 2 To LastRow2
                PerChange = ws.Cells(m, 11)
                SValueSum = ws.Cells(m, 12)
                Ticker = ws.Cells(m, 9)
                
                If PerChange > PerChangeMax Then
                    PerChangeMax = PerChange
                    ws.Cells(2, 16) = PerChangeMax
                    ws.Cells(2, 15) = Ticker
                ElseIf PerChange < PerChangeMin Then
                    PerChangeMin = PerChange
                    ws.Cells(3, 16) = PerChangeMin
                    ws.Cells(3, 15) = Ticker
                ElseIf SValueSum > SValueSumMax Then
                    SValueSumMax = SValueSum
                    ws.Cells(4, 16) = SValueSumMax
                    ws.Cells(4, 15) = Ticker
                End If
        Next m
        
        '-----------------------------------------------------------------
        'formatting
        For k = 2 To LastRow2
            If ws.Cells(k, 10).Value > 0 Then
                ws.Cells(k, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(k, 10).Value < 0 Then
                ws.Cells(k, 10).Interior.ColorIndex = 3
            End If
            
        Next k
        
        For l = 2 To LastRow2
            ws.Cells(l, 11).NumberFormat = "0.00%"
        Next l
        
        '-------------------------------------------------------------------
        'original table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Pecent Change"
        ws.Cells(1, 12).Value = "Total Stock Value"
        
        'secondary table
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
        '--------------------------------------------------------------------
        'Reset Min/Max Variables
        PerChangeMax = 0
        PerChangeMin = 0
        SValueSumMax = 0
        
    
    
    Next ws
    
    MsgBox ("Completed")
    
End Sub
