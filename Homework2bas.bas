Attribute VB_Name = "Module1"

Option Explicit
Sub stock_1()
    'Remember Excel start at 1
    'Remember VBA starts at 0
    Dim ws As Worksheet
    
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
    
        Dim LastRow As Long, i As Long, CellValue As String
        Dim Letters As String

        With ActiveSheet  'Get last row cell value to help set up array below
                LastRow = Cells(.Rows.Count, "A").End(xlUp).Row
        End With
  
    
        Dim Max_Rows As Integer  'Get Max number of unique Ticker Symbols
        Max_Rows = 1

        Letters = Cells(2, 1).Value
 
       For i = 2 To LastRow
                CellValue = Cells(i, 1).Value
                    If Letters <> CellValue Then
                       Max_Rows = Max_Rows + 1
                    Else
                      Max_Rows = Max_Rows + 0
                   End If
                     Letters = CellValue
      Next
                

'Define an Array of variant to store columns we want to sum up and process for later
ReDim StoreData(Max_Rows, 10) As Variant
Dim I_Row As Integer
Dim I_Col As Integer
Dim First_Entry As Integer
Dim Percent_Change As Double


'Initialize row and column counters and any flags
I_Row = 0
I_Col = 0
First_Entry = 0 'Store first cell value


    
     CellValue = Cells(2, 1).Value
     Letters = CellValue


   For i = 3 To LastRow
       CellValue = Cells(i, 1).Value
      If Letters <> CellValue Then
        'When a change is detected grab the previous records values
            
            StoreData(I_Row, 4) = Cells(i - 1, 2).Value 'Ending Date
            StoreData(I_Row, 5) = Cells(i - 1, 6).Value 'Closing Price
            StoreData(I_Row, 3) = StoreData(I_Row, 3) + Cells(i - 1, 7).Value   'Volume
            'In_Open = Cells(i - 1, 1).Value
            I_Row = I_Row + 1
            First_Entry = 0
        Else
          'When no change is detected grab the current record once and the frist record only once
          
            If First_Entry = 0 Then
            StoreData(I_Row, 0) = Cells(i - 1, 1).Value
            StoreData(I_Row, 1) = Cells(i - 1, 2).Value 'Begging Date
            StoreData(I_Row, 2) = Cells(i - 1, 3).Value 'Open Price
            StoreData(I_Row, 3) = Cells(i - 1, 7).Value 'Volume
            First_Entry = First_Entry + 1
            Else
            'Ensure all the Volumes are accumulated
            StoreData(I_Row, 3) = StoreData(I_Row, 3) + Cells(i - 1, 7).Value   'Volume
            End If
            
        End If
        Letters = CellValue
   Next
   
   'output results to spreadhsheet
   'Format Top Line
   Cells(1, 9).Value = "Ticker"
   Cells(1, 10).Value = "Yearly Change"
   Cells(1, 11).Value = "Percent Change"
   Cells(1, 12).Value = "Total Stock Volume"
   


'Process through array to display columns output as needed
 For i = 1 To Max_Rows
   Cells(i + 1, 9).Value = StoreData(i - 1, 0)
   Cells(i + 1, 10).Value = CDbl(StoreData(i - 1, 5)) - CDbl(StoreData(i - 1, 2))
   
   If StoreData(i - 1, 2) = 0 Then  'handle 0 denominator to avoid divide by zero 0 error
      Percent_Change = 0
  Else
      Percent_Change = (CDbl(StoreData(i - 1, 5)) / CDbl(StoreData(i - 1, 2))) - 1
   End If
 
   
   Cells(i + 1, 11).Value = Format(Percent_Change, "Percent")
   
   
   Cells(i + 1, 12).Value = StoreData(i - 1, 3)

   If Percent_Change < 0 Then  'Change format color of yearly change to red is negative and green if positive
            Cells(i + 1, 11).Interior.ColorIndex = 3 ' 5 indicates Blue Color, 3 = red
   Else
            Cells(i + 1, 11).Interior.ColorIndex = 4 ' 5 indicates Blue Color,  4= green
   End If

   Next i '
   'output results to spreadhsheet
   'Format Top Line
   Cells(1, 15).Value = ""
   Cells(1, 16).Value = "Ticker"
   Cells(1, 17).Value = "Value"
   Cells(2, 15).Value = "Greatest % Increase"
   Cells(3, 15).Value = "Greatest % Decrease"
   Cells(4, 15).Value = "Greatest Total Volume"



'Calculate final few values
Dim Max_Inc As Double
Dim Min_Inc As Double
Dim Max_VComp As Double
Dim Max_Vol As Double
Dim Ticker_I As String
Dim Ticker_J As String
Dim Ticker_V As String
Dim t As Integer

Max_Inc = 0
Percent_Change = 0
Max_VComp = 0
Max_Vol = 0

For t = 1 To Max_Rows

   If StoreData(t - 1, 2) = 0 Then
      Percent_Change = 0
Else
   

   Percent_Change = (CDbl(StoreData(t - 1, 5)) / CDbl(StoreData(t - 1, 2))) - 1
   End If
 

'Percent_Change = (CDbl(StoreData(t - 1, 5)) / CDbl(StoreData(t - 1, 2))) - 1

    If Percent_Change >= Max_Inc Then
        Max_Inc = Percent_Change
        Ticker_I = StoreData(t - 1, 0)
    Else
       Min_Inc = Percent_Change
       Ticker_J = StoreData(t - 1, 0)
    End If
Next t

'For Max Volume
For t = 1 To Max_Rows
Max_VComp = StoreData(t - 1, 3)

    If Max_VComp >= Max_Vol Then
        Max_Vol = Max_VComp
        Ticker_V = StoreData(t - 1, 0)

    End If
Next t

 Cells(2, 16).Value = Ticker_I
 Cells(2, 17).Value = Format(Max_Inc, "Percent")
 Cells(3, 16).Value = Ticker_J
 Cells(3, 17).Value = Format(Min_Inc, "Percent")
  Cells(4, 16).Value = Ticker_V
 Cells(4, 17).Value = Max_Vol


Next


End Sub

