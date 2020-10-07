Attribute VB_Name = "Module1"

Sub main()

 ' Declare Current as a worksheet object variable.
         Dim Current As Worksheet

         Dim thisarray As Variant
         Dim lastRow As Long
         Dim uniques As Collection
         Dim source As Range
         Dim tickerarray As Variant
  
         ' Loop through all of the worksheets in the active workbook.
         For Each Current In Worksheets
            Dim i As Integer: i = 0
   
             lastRow = Cells(Rows.Count, "G").End(xlUp).Row
             
              Set tickersource = Current.Range("A2:A" & lastRow)
              Set uniques = GetUniqueValues(tickersource.Value)
              'tickerarray = collectionToArray(uniques)
              thisarray = Current.Range("A1:G" & lastRow).Value
            
              counter = 1
             
            'start working on filling the final array, that i will use for posting values into worksheet
             Dim final() As String
             ReDim final(1 To 4, uniques.Count)
             
             final(1, 0) = "Ticker"
             final(2, 0) = "Yearly Change"
             final(3, 0) = "Percent Change"
             final(4, 0) = "Total Stock Volume"
             
             
            'fill the ticker part of the final array
             Dim ticker As Variant
             For Each ticker In uniques
                final(1, counter) = ticker
                counter = counter + 1
             Next ticker
            
            'reset for looping through the thisarray
            'thisarray is 1-based, and first row is labels
             counter = 2
             
             'Dim yearly As Integer
             Dim final_counter As Integer
             final_counter = 1
             
             'dim temp variables for calculation
             Dim sum_vol As Double
             sum_vol = 0
             
             Dim price_open As Double
             price_open = 0
             
             Dim price_close As Double
             price_close = 0
             
             While (counter) <= UBound(thisarray)
               price_open = thisarray(counter, 3)
               
               Do While thisarray(counter, 1) = final(1, final_counter)
                    
                    'calculate volume
                    sum_vol = sum_vol + thisarray(counter, 7)
                    price_close = thisarray(counter, 6)
                
                If counter = UBound(thisarray) Then
                    Exit Do
                Else
                    counter = counter + 1
                End If
                
               Loop
               
               final(4, final_counter) = CStr(sum_vol)
               sum_vol = 0
               
               'calculate change in price
               final(2, final_counter) = CStr(price_close - price_open)
               
               'calculate percentage change in price
               If (price_open <> 0) Then
                    final(3, final_counter) = CStr(Round(((price_close - price_open) / price_open * 100), 2))
               Else
                final(3, final_counter) = 0
               End If
               
               final_counter = final_counter + 1
               If (final_counter > UBound(final, 2)) Or (counter > UBound(thisarray)) Then GoTo nextsheet
               
             Wend
                
nextsheet:
                counter = 1
             
             'post values from final into worksheet
             Call PrintArray(final, Current.Name)
             
             'clear final array
             ReDim final(0, 0)
         Next
   
End Sub

Public Function PrintArray(final As Variant, sheetname As String)

Dim finalarray As Variant
finalarray = WorksheetFunction.Transpose(final)

Worksheets(sheetname).Range("I1").Resize(UBound(finalarray, 1), UBound(finalarray, 2)).Value = finalarray

'Set Summary = Range("I1:").Resize(Transpose(UBound(finalarray, 1)))
' Summary.Value = finalarray

 
End Function

Public Function GetUniqueValues(ByVal values As Variant) As Collection
    Dim result As Collection
    Dim cellValue As Variant
    Dim cellValueTrimmed As String

    Set result = New Collection
    Set GetUniqueValues = result

    On Error Resume Next

    For Each cellValue In values
        cellValueTrimmed = Trim(cellValue)
        If cellValueTrimmed = "" Then GoTo NextValue
        result.Add cellValueTrimmed, cellValueTrimmed
NextValue:
    Next cellValue

    On Error GoTo 0
End Function

Public Function collectionToArray(c As Collection) As Variant()
    Dim a() As Variant: ReDim a(0 To c.Count - 1)
    Dim i As Integer
    For i = 1 To c.Count
        a(i - 1) = c.Item(i)
    Next
    collectionToArray = a
End Function
  
  

