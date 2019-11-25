Attribute VB_Name = "Module1"
Option Explicit

Sub A()

Dim ws As Worksheet



Dim i As Long
Dim j As Integer
Dim last_row As Long

Dim year_end As Double
Dim year_start As Double
Dim per_change As Double
Dim year_change As Double
Dim stock_volume As Double

Dim gr_incr_per As Double
Dim gr_decr_per As Double
Dim gr_incr_ticker As String
Dim gr_decr_ticker As String
Dim gr_volume As Double
Dim gr_volume_ticker As String



For Each ws In Worksheets

ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Wolume"

ws.Cells(1, 16) = "Ticker"
ws.Cells(1, 17) = "Value"
ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(4, 15) = "Greatest Total Volume"


'calculate last row
last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row

year_start = ws.Cells(2, 3)
j = 2
stock_volume = 0

gr_incr_per = 0
gr_decr_per = 0
gr_volume = 0


For i = 2 To last_row

    stock_volume = stock_volume + ws.Cells(i, 7)
    
    If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
        ws.Cells(j, 9) = ws.Cells(i, 1)
        
        'calculate year change
        year_end = ws.Cells(i, 6)
        year_change = year_end - year_start
        ws.Cells(j, 10) = year_change
        
        'set cells color for year change
        If year_change > 0 Then
                ws.Cells(j, 10).Interior.Color = vbGreen
        Else
                ws.Cells(j, 10).Interior.Color = vbRed
        
        End If
        
        'calculate yearly percentage change
        If year_start <> 0 Then
            per_change = (year_end - year_start) / year_start
        Else
            per_change = 0
        End If
        
        ws.Cells(j, 11) = per_change
        ws.Cells(j, 11).NumberFormat = "0.00%"
        
        If per_change > gr_incr_per Then
            gr_incr_per = per_change
            gr_incr_ticker = ws.Cells(i, 1)
        End If
        
            
        If per_change < gr_decr_per Then
            gr_decr_per = per_change
            gr_decr_ticker = ws.Cells(i, 1)
        End If
        
        If stock_volume > gr_volume Then
            gr_volume = stock_volume
            gr_volume_ticker = ws.Cells(i, 1)
        End If
        
            
        
       'reset year_start and stock_volume
        year_start = ws.Cells(i + 1, 3)
        
        ws.Cells(j, 12) = stock_volume
        stock_volume = 0
        
        j = j + 1
    End If
    

Next i

    ws.Cells(2, 17) = gr_incr_per
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(2, 16) = gr_incr_ticker
    
    ws.Cells(3, 17) = gr_decr_per
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16) = gr_decr_ticker
    
    ws.Cells(4, 17) = gr_volume
    ws.Cells(4, 16) = gr_volume_ticker
    
Next ws

End Sub


