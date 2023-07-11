Attribute VB_Name = "Module1"
Option Explicit    'to prevent latent errors
Sub stock_ticker():
  'Color Red
  Const COLOR_RED As Integer = 3
  Const COLOR_GREEN As Integer = 4
  ' Set an initial variable for holding the ticker
  Dim Ticker As String
  
  'define the input_row counter
  Dim input_row As Long
  'enable looping through multiple worksheets
  Dim WS As Worksheet
  ' Set an initial variable for Yearly open,Yearly close, Yearly change and color...
  Dim Yearly_Open As Double
  Dim Yearly_Close As Double
  Dim Yearly_Change As Double
  Dim Yearly_Color As Integer
   ' Set an initial variable for Percent change, called here Change_Ratio
  Dim Change_Ratio As Double
  Dim Ratio_Color As Integer
   ' Set an initial variable for Total Stock Volum
  Dim Total_Volume As LongLong
  ' Keep track of the location for each stock ticker and values in the summary table
  Dim Summary_Table_Row As Integer
  Dim Max_Incr_Ticker As String
  Dim Min_Decr_Ticker As String
  Dim Max_Vol_Ticker As String
  Dim Max_Incr_Value As Double
  Dim Min_Decr_Value As Double
  Dim Max_Vol_Value As LongLong
  
  For Each WS In Worksheets
    WS.Activate

      'setup for first stock
      Summary_Table_Row = 2
      Total_Volume = 0
      Yearly_Open = Cells(2, 3).Value
      Max_Incr_Value = -999
      Min_Decr_Value = 999
      Max_Vol_Value = -999
      ' Loop through all stocks
      For input_row = 2 To Cells(Rows.Count, "A").End(xlUp).Row
        ' Set the Ticker
          Ticker = Cells(input_row, 1).Value
          Total_Volume = Total_Volume + Cells(input_row, 7).Value
          
        ' last row of current stock
        If Cells(input_row + 1, 1).Value <> Ticker Then
          'Input
          Yearly_Close = Cells(input_row, 6).Value
          
          'Calculations
          Yearly_Change = (Yearly_Close - Yearly_Open)
          If Yearly_Change >= 0 Then
            Yearly_Color = COLOR_GREEN
          Else
            Yearly_Color = COLOR_RED
          End If
          
          Change_Ratio = Yearly_Change / Yearly_Open
          If Change_Ratio >= 0 Then
            Ratio_Color = COLOR_GREEN
          Else
            Ratio_Color = COLOR_RED
          End If
   'Summary Table Calculations
          If Change_Ratio > Max_Incr_Value Then
            Max_Incr_Value = Change_Ratio
            Max_Incr_Ticker = Ticker
          End If
          If Change_Ratio < Min_Decr_Value Then
            Min_Decr_Value = Change_Ratio
            Min_Decr_Ticker = Ticker
          End If
          If Total_Volume > Max_Vol_Value Then
            Max_Vol_Value = Total_Volume
            Max_Vol_Ticker = Ticker
          End If
          
          'Output
          Range("I" & Summary_Table_Row).Value = Ticker
          Range("J" & Summary_Table_Row).Value = Yearly_Change
          Range("J" & Summary_Table_Row).Interior.ColorIndex = Yearly_Color
          Range("K" & Summary_Table_Row).Value = FormatPercent(Change_Ratio)
          Range("K" & Summary_Table_Row).Interior.ColorIndex = Ratio_Color
          Range("L" & Summary_Table_Row).Value = Total_Volume
          
          ' setup for next stock
          Summary_Table_Row = Summary_Table_Row + 1
          Total_Volume = 0
          Yearly_Open = Cells(input_row + 1, 3).Value
    
        End If
    
      Next input_row
      MsgBox input_row
    'Header Information
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    ' second summary table Output
    Range("P" & 2).Value = Max_Incr_Ticker
    Range("Q" & 2).Value = FormatPercent(Max_Incr_Value)
    Range("P" & 3).Value = Min_Decr_Ticker
    Range("Q" & 3).Value = FormatPercent(Min_Decr_Value)
    Range("P" & 4).Value = Max_Vol_Ticker
    Range("Q" & 4).Value = Max_Vol_Value
    
    Next WS
    
    MsgBox "done"
End Sub

