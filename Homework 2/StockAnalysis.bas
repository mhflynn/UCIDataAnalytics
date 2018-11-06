Attribute VB_Name = "Module1"
'------------------------------------------------------------------------------
'
' For each worksheet in a file, this function will process a table
' containing stock price and volume data by ticker symbol. Each 
' worksheet is assumed to conatain data for a single year. 
' A summary table will be created on the same spreadsheet, to the right
' of the original table, containing ticker symbol and total volume.
' Additionally a post processing step will be run on the summary table,
'   - Conditionally color cells with percent change, green for positive
'     gains, red for negative gains for the year
'   - Output an additional table with values for Ticker symbols with
'     least gain, highest gain and highest volume
'
' Assumptions for input table :
'   - The first cell of the input table is "A1"
'   - The input spans 7 columns "A" - "G"
'   - The table spans a variable number of rows
'   - The end of the table is the first row with empty cell in column "A"
'   - The first row in the table contains headers as below
'       |<ticker>| <date> | <open> | <high> | <low> | <close> | <volume> |
'       The content of the header row may be checked, but it is assumed that
'       that data in the respective columns contains valid values
'   - <ticker> symbols are text of variable length
'   - <ticker> symbols are grouped in consecutive rows for each unique symbol
'   - A new <ticker> symbol signals processing for current symbol is complete
'   - <date> values are in chronological order
'       - First row for each <ticker> symbol contains open price for the year
'       - Last row for each <ticker> symbols contains closing price for the year
'   - <date> values have format YYYYMMDD. e.g. "20140131" for Jan 31, 2014
'   - <open>, <high>, <low> and <close> are "float" values
'   - <volume> values are of "long" type
'   - All dates in the table are for the same year
'   - A value of zero for first open price for a ticker symbol assumes there
'     is no volume, hence no gain for the associated stock.
'
' Assumption for output table :
'   - There is no other data in the spreadsheet beyond input table boundaries
'   - Columns for output data will be cleared without loss of data
'   - The first cell of the output table will be "I1"
'   - The output spans columns "I" through "P"
'   - The first row of output will contains headers as below
'       | Ticker | Yearly Change | Percent Change | Total Volume |
'   - Ticker column will contain the <ticker> symbols for input table
'   - Yearly Change will contain "float" value for dollar amount
'   - Percent Change will contain "float" value for percent value
'   - Total Volume will contain "long" value for total shares traded
'
' Assumptions for post processing output :
'   - The first cell of the output table will be "N1"
'   - The output spans columns "N" through "P" worst case
'   - The post processing output will be 3 x 4 area from "N1" containing
'     labels and the values for least/most gain and highest volume
'
' TODO : Mitigate the assumptions from above :
'   - Handle cases with zero for first opening price, check for subsequent
'      non-zero price values and use for opening price
'   - Support unordered table, ticker symbols not grouped, transactions
'      not in chronalogical order
'   - Support variable start column for input table, summary and post processSub StockAnalysis()
'
'------------------------------------------------------------------------------
'

'------------------------------------------------------------------------------
'
Sub StockAnalysis()

'Initialize the Status Bar to indicate progress
SaveBarStat = Application.DisplayStatusBar
Application.DisplayStatusBar = True

'Loop for each Worksheet in the active Workbook
For Each Sheet In Worksheets

    Sheet.Activate          ' Activate the Worksheet
    Call Analyze             ' Process the stock data input table
    Call PostProcess       ' Summarize the analysis output
    
Next Sheet

'Restore Status Bar state
Application.DisplayStatusBar = SaveBarStat
Application.StatusBar = False

End Sub

'
'------------------------------------------------------------------------------
'
Private Sub Analyze()

Dim Ticker, NextTicker As String
Dim OpenValue, GainValue As Double
Dim CurInRow, CurOutRow As Integer
Dim Volume, VolumeNext As Long

Application.StatusBar = "Start Analyze ..."
Columns("I:P").Clear

'Initialize variables for first Ticker symbol

CurOutRow = 2                                                 ' Index for first output row

CurInRow = 2                                                  ' Index for first input row
Ticker = Cells(CurInRow, "A").Value                           ' Initial Ticker symbol
OpenValue = Cells(CurInRow, "C").Value                        ' Initial opening price
GainValue = Cells(CurInRow, "F").Value                        ' Initial closing price
Volume = Cells(CurInRow, "G").Value                           ' Initial volume
CurInRow = CurInRow + 1                                       ' Move to the next input row

Application.StatusBar = "Analyze ..."

If (Ticker <> "") Then

'   If table is not empty, loop through each row of the input table
    Do While (True)
    
'       Process data for Ticker symbol in the current row
    
        TickerNext = Cells(CurInRow, "A").Value               ' Get ticker symbol from the current row
        
        If (Ticker = TickerNext) Then                         ' Check if new Ticker symbol encountered
        
'           No change in ticker symbol, continue to accumulate data for current symbol
    
            Volume = Volume + Cells(CurInRow, "G").Value      ' Accumulate volume
            GainValue = Cells(CurInRow, "F")                  ' Save latest closing price
    
        Else
            
'           New Ticker symbol found, this row starts data for new symbol
'           Finalize and output calculation for current symbol
    
            Cells(CurOutRow, "I").Value = Ticker              ' Ticker symbol to output table
            GainValue = GainValue - OpenValue                 ' Calculate yearly gain
            Cells(CurOutRow, "J").Value = GainValue           ' Yearly gain to the output table
            Cells(CurOutRow, "L").Value = Volume              ' Yearly volume to the output table
            
            If (OpenValue = 0) Then                           ' Calulate % gain, check for divide by 0
                Cells(CurOutRow, "K").Value = 0
            Else
                Cells(CurOutRow, "K").Value = GainValue / OpenValue
            End If
     
            CurOutRow = CurOutRow + 1                         ' Position output for the next row
     
            If (TickerNext = "") Then                         ' Check if reach end of input
                Exit Do                                       ' At the end, exit row loop
                
'           Initialize variables for new symbol
     
            Else
                Ticker = TickerNext                           ' Establish the current Ticker symbol
                OpenValue = Cells(CurInRow, "C").Value        ' Initialize the open price for the year
                GainValue = Cells(CurInRow, "F").Value        ' Capture latest closing price
                Volume = Cells(CurInRow, "G").Value           ' Initial volume

                'Application.StatusBar = "Analyze : " & CurOutRow & "    " & Ticker
            End If
        End If
        CurInRow = CurInRow + 1                               ' Move to the next input row
    
    Loop

    Call FormatAnalyzeOutputArea                              ' Format the output table area
End If

Cells(1, "I").Select
Application.StatusBar = "Analyze Complete"

End Sub
'
'------------------------------------------------------------------------------
'
Private Sub PostProcess()

'Initial index for first row
CurInRow = 2

'Initial Ticker symbols for the stats
MinSym = Cells(CurInRow, "I")
MaxSym = MinSym
VolSym = MinSym

'Initial values for the stats
MinPct = Cells(CurInRow, "K")
MaxPct = MinPct
MaxVol = Cells(CurInRow, "L")

Application.StatusBar = "Post Processing ..."

'Loop through each row of the analysis output, collect the stats

Do While (True)
    
    If (Cells(CurInRow, "I") = "") Then Exit Do          ' If reach the end of the analysis output, exit the loop
    
    If (MinPct > Cells(CurInRow, "K")) Then              ' Check for new minimum percent change
        MinPct = Cells(CurInRow, "K")
        MinSym = Cells(CurInRow, "I")
    End If
    
    If (MaxPct < Cells(CurInRow, "K")) Then              ' Check for new maximum percent change
        MaxPct = Cells(CurInRow, "K")
        MaxSym = Cells(CurInRow, "I")
    End If
    
    If (MaxVol < Cells(CurInRow, "L")) Then              ' Check for new maximum volume
        MaxVol = Cells(CurInRow, "L")
        VolSym = Cells(CurInRow, "I")
    End If
    
    If (Cells(CurInRow, "J") < 0) Then                   ' Format the percentage gain value
        Cells(CurInRow, "J").Interior.ColorIndex = 3     ' Negative gain, fill cell Red
    Else
        Cells(CurInRow, "J").Interior.ColorIndex = 43    ' Positive gin, fill cell Green
    End If
    
    CurInRow = CurInRow + 1                              ' Move to next row

Loop

'If analysis output table was not empty, format the summary output area
If (MinSym <> "") Then

'   Stat values to the output area
    Cells(2, "O").Value = MaxSym
    Cells(2, "P").Value = MaxPct
    Cells(3, "O").Value = MinSym
    Cells(3, "P").Value = MinPct
    Cells(4, "O").Value = VolSym
    Cells(4, "P").Value = MaxVol
    
    Call FormatPostProcessingOutputArea
    
End If

Application.StatusBar = "Post Processing Complete"

End Sub
'
'------------------------------------------------------------------------------
'
Private Sub FormatAnalyzeOutputArea()
'
' Macro5 Macro
'

'
'   Setup and format the output area

    Cells(1, "I").Value = "  Ticker  "
    Cells(1, "J").Value = "  Yearly Change  "
    Cells(1, "K").Value = "  Percent Change  "
    Cells(1, "L").Value = "  Total Volume  "
    
    Columns("I:L").Select
    Selection.Columns.AutoFit
    Columns("I:I").Select
    Selection.NumberFormat = "General"
    Columns("J:J").Select
    Selection.NumberFormat = "0.00"
    Columns("K:K").Select
    Selection.NumberFormat = "0.00%"
    Columns("L:L").Select
    Selection.NumberFormat = "0"
    Columns("J:J").Select
    Columns("I:L").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
    Cells(1, "I").Select
End Sub

'
'------------------------------------------------------------------------------
'
Private Sub FormatPostProcessingOutputArea()
    
'   Stat header data to the output area
    Cells(1, "O").Value = "  Ticker  "
    Cells(1, "P").Value = "  Amount  "
    Cells(2, "N").Value = "Greatest % Increase"
    Cells(3, "N").Value = "Greatest % Decrease"
    Cells(4, "N").Value = "Greatest Volume Amount"
    
'   Format the post processing output area
    Range("N1", "N4").HorizontalAlignment = xlLeft
    Range("O1", "P4").HorizontalAlignment = xlCenter
    Range("P2", "P3").NumberFormat = "0.00%"
    Range("P4", "P4").NumberFormat = "0"
    Range("N1", "N4").HorizontalAlignment = xlLeft
    Columns("N:P").Columns.AutoFit
End Sub

