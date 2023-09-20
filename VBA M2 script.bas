Attribute VB_Name = "Module1"
Option Explicit
Public dblP_Increase_Tic As String
Public dblP_Decrease_Tic As String
Public dblT_Volume_Tic As String

Public dblP_Increase_Val As Double
Public dblP_Decrease_Val As Double
Public dblT_Volume_Val As Double


Sub Main()
'Declaration of global variables
'Declaration of local variables
Dim strP_Ticker As String   'stores the previous Ticker value
Dim strC_Ticker As String   'stores the current Ticker value
Dim dblO_Price As Double    'stores the opening price of the year
Dim dblC_Price As Double    'stores the closing price of the year
Dim dblD_Price As Double    'stores the difference of the opening and closing price of the year
Dim dblP_Change As Double   'stores the percentage change of the opening and closing price of the year
Dim lngT_Vol As Double      'stores the total stock volume of the year
Dim iRow As Long
Dim EOYRow As Integer
Dim int_WkshtIndex As Integer
Dim i As Integer

'Total number of worksheets
int_WkshtIndex = ActiveWorkbook.Sheets.Count


'This will loop through all the worksheets and perform necessary actions/calculations as per requirement
For i = 1 To int_WkshtIndex
    'Initialize the variables
    iRow = 2
    EOYRow = 2
    dblD_Price = 0
    dblP_Change = 0
    lngT_Vol = 0
    
    dblP_Increase_Tic = ""
    dblP_Decrease_Tic = ""
    dblT_Volume_Tic = ""
    
    dblP_Increase_Val = 0
    dblP_Decrease_Val = 0
    dblT_Volume_Val = 0
    
    ActiveWorkbook.Sheets(i).Select
    FillHeaderNames i
    
    
    dblO_Price = Sheets(i).Cells(iRow, 3).Text
    Do
        strC_Ticker = Sheets(i).Cells(iRow, 1).Text
        If strC_Ticker <> "" Then lngT_Vol = lngT_Vol + Sheets(i).Cells(iRow, 7).Text
        
        strP_Ticker = strC_Ticker
    
        iRow = iRow + 1
        strC_Ticker = Sheets(i).Cells(iRow, 1).Text
        
        If strC_Ticker <> strP_Ticker Then
            dblC_Price = Sheets(i).Cells(iRow - 1, 6).Text
            dblD_Price = dblC_Price - dblO_Price
            dblP_Change = (dblC_Price - dblO_Price) / dblO_Price
            
            FillEOY i, EOYRow, strP_Ticker, dblD_Price, dblP_Change, lngT_Vol
            EOYRow = EOYRow + 1
        
            If strC_Ticker <> "" Then
                dblO_Price = Sheets(i).Cells(iRow, 3).Text
                lngT_Vol = 0
            End If
        End If
    
    Loop While strC_Ticker <> ""
    
    ConditionalFormat i
    FillGreatest i
Next i
    
End Sub

Sub FillHeaderNames(wkshtIndex As Integer)
    
    With Sheets(wkshtIndex)
        '.Activate
        
        'Ticker
        .Cells(1, 9) = "Ticker"
        
        'Yearly Change
        .Cells(1, 10) = "Yearly Change"
        .Columns(10).NumberFormat = "0.00"
        
        'Percent Change
        .Cells(1, 11) = "Percent Change"
        .Columns(11).Style = "Percent"
        .Columns(11).NumberFormat = "0.00%"

        
        'Total Stock Volume
        .Cells(1, 12) = "Total Stock Volume"
        .Columns(12).NumberFormat = "0"
        
        'Greatest - Column
        .Cells(1, 16) = "Ticker"
        .Cells(1, 17) = "Value"
        
        'Greatest - Row
        .Cells(2, 15) = "Greatest % Increase"
        .Cells(3, 15) = "Greatest % Decrease"
        .Cells(4, 15) = "Greatest Total Volume"
    End With
    
End Sub

Sub ConditionalFormat(wkshtIndex As Integer)

Dim iCol As Integer

For iCol = 10 To 11

    With Sheets(wkshtIndex)
        '.Activate
        
        .Columns(iCol).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        .Columns(iCol).FormatConditions(.Columns(iCol).FormatConditions.Count).SetFirstPriority
        With .Columns(iCol).FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13561798
            .TintAndShade = 0
        End With
        
        .Columns(iCol).FormatConditions(1).StopIfTrue = False
                
        .Columns(iCol).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
        .Columns(iCol).FormatConditions(.Columns(iCol).FormatConditions.Count).SetFirstPriority
        With .Columns(iCol).FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        
        .Columns(iCol).FormatConditions(1).StopIfTrue = False
        
        .Cells(1, iCol).FormatConditions.Delete
    End With
    
Next iCol
        
End Sub

Sub FillEOY(wkshtIndex As Integer, EOYRow As Integer, strP_Ticker As String, dblD_Price As Double, dblP_Change As Double, lngT_Vol As Double)

    With Sheets(wkshtIndex)
        '.Activate
        
        .Cells(EOYRow, 9) = strP_Ticker
        .Cells(EOYRow, 10) = dblD_Price
        .Cells(EOYRow, 11) = dblP_Change
        .Cells(EOYRow, 12) = lngT_Vol
        

        If dblP_Increase_Val < dblP_Change Then
            dblP_Increase_Val = dblP_Change
            dblP_Increase_Tic = strP_Ticker
        End If
        
        If dblP_Decrease_Val > dblP_Change Then
            dblP_Decrease_Val = dblP_Change
            dblP_Decrease_Tic = strP_Ticker
        End If
    
        If dblT_Volume_Val < lngT_Vol Then
            dblT_Volume_Val = lngT_Vol
            dblT_Volume_Tic = strP_Ticker
        End If
    End With

End Sub

Sub FillGreatest(wkshtIndex As Integer)

    With Sheets(wkshtIndex)
        '.Activate
        
        'Greatest % Increase
        .Cells(2, 16) = dblP_Increase_Tic
        .Cells(2, 17) = dblP_Increase_Val
        .Cells(2, 17).Style = "Percent"
        .Cells(2, 17).NumberFormat = "0.00%"
        
        'Greatest % Decrease
        .Cells(3, 16) = dblP_Decrease_Tic
        .Cells(3, 17) = dblP_Decrease_Val
        .Cells(3, 17).Style = "Percent"
        .Cells(3, 17).NumberFormat = "0.00%"

        'Greatest Total Volume
        .Cells(4, 16) = dblT_Volume_Tic
        .Cells(4, 17) = dblT_Volume_Val
        '.Cells(4, 17).NumberFormat = "0"


    End With
    
End Sub
