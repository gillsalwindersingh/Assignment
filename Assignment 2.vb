Sub AllSheets()
Dim ws As Worksheet
Dim rowcount As Long
rowcount = 2
For Each ws In Worksheets
        ws.Activate
    'put your row insert code here
        For i = 2 To 70926
                Row = 1
                totalvolume = totalvolume + Cells(i, 7).Value
                
                ticker = Cells(i, 1).Value
                If Cells(i + Row, 1).Value <> Cells(i, 1).Value Then
                rowcount = rowcount + 1
                Range("I" & rowcount).Value = ticker
                Range("L" & rowcount).Value = totalvolume
                
                
                                
          End If
                
         Next i
    
    Next ws
    
End Sub