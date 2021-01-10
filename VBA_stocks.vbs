Attribute VB_Name = "Module1"
Sub Stock()
    
    Dim ws As Worksheet

    For Each ws In Worksheets
        ws.Activate
        
        Dim ticker As String
        Dim change As Double
        Dim total As Double
        Dim opn As Long

        total = 0
        change = 0
        opn = 2
        summary = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        For i = 2 To lastRow
            
            If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
                
                ticker = Cells(i, "A").Value
                total = total + Cells(i, "G").Value
    
                If total = 0 Then
                    Cells(summary, "I").Value = Cells(i, "A").Value
                    Cells(summary, "J").Value = 0
                    Cells(summary, "K").Value = 0 & "%"
                    Cells(summary, "L").Value = 0
            
                Else
                    If Cells(opn, "C") = 0 Then
                        For smry = opn To i
                            If Cells(smry, "C").Value <> 0 Then
                                opn = smry
                                Exit For
                            End If
                        Next smry
                    End If
                    
                    change = (Cells(i, "F").Value - Cells(opn, "C").Value)
                    percent = Round((change / Cells(opn, "C") * 100), 2)
                    
                    Cells(summary, "I").Value = ticker
                    Cells(summary, "J").Value = change
                    Cells(summary, "K").Value = "%" & percent
                    Cells(summary, "L").Value = total
                   
                    If change > 0 Then
                            Cells(summary, "J").Interior.ColorIndex = 4
                    ElseIf change < 0 Then
                            Cells(summary, "J").Interior.ColorIndex = 3
                    Else
                            Cells(summary, "J").Interior.ColorIndex = 2
                    End If
                    
                End If
                   summary = summary + 1
                   opn = i + 1
                   total = 0
                   change = 0
           Else
                total = total + Cells(i, "G").Value
           End If
        Next i
        
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        Cells(1, "I").Font.Bold = True
        Cells(1, "J").Font.Bold = True
        Cells(1, "K").Font.Bold = True
        Cells(1, "L").Font.Bold = True
               
    Next ws
    
End Sub
