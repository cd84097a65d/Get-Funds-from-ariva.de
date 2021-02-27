Attribute VB_Name = "Funds"
Option Explicit

Const startingLine = 3      ' start line at sheet "Funds"
Const clmSortingIndex = 8   ' column H, because index is at cell H1

Const clmWKN = 3
Const clmFavorites = 4
Const clmChangeOfPosition = 5
Const clmStockOrFund = 6
Const clmCountry = 7
Const clmSector = 8
Const clmBenchmark = 9
Const clmCurrency = 10
Const clmUrl_ariva = 11
Const clmUrl_finanzen = 12
Const clm_3_months = 13
Const clm_6_months = 14
Const clm_1_year = 15
Const clm_3_years = 16
Const clm_5_years = 17
Const clmDate = 18
Const clmPrice = 19
Const clmAlpha = 20
Const clmBeta = 21
Const clmSharpeRatio = 22
Const clmVolatility = 23
Const clmTrackingError = 24
Const clmCorrelation = 25
Const clmSkewness = 26
Const clmKurtosis = 27
Const clmSortinoRatio = 28
Const clmInformationRatio = 29
Const clmR2 = 30
Const clmTreynorRatio = 31

Dim wsFunds As Worksheet

Sub Update_Funds()
    Dim url$, wkn$, price$, currency_$, development$(5), country$, sector$, benchmark$
    Dim alpha$, beta$, sharpeRatio$, volatility$, trackingError$, correlation$
    Dim skewness$, kurtosis$, sortinoRatio$, informationRatio$, r2$, treynorRatio$
    Dim url_Finanzen$, decimalSeparator$
    Dim wkns$()
    Dim updated() As Boolean
    Dim i%, j%
    Dim favorites As Boolean
    
    decimalSeparator = Application.decimalSeparator
    Application.decimalSeparator = ","
    
    Set wsFunds = Worksheets("Funds")
    favorites = wsFunds.Cells(1, 2)
    Selenium_ariva.cookiesAccepted = False
    Selenium_Finanzen.cookiesAccepted = False
    
    ' sorting (3m)
    wsFunds.Cells(1, clmSortingIndex) = 1
    Sorting_Funds
    
    i = startingLine
    While wsFunds.Cells(i, clmWKN) <> ""
        wkn = wsFunds.Cells(i, clmWKN)
        url = wsFunds.Cells(i, clmUrl_ariva)
        
        ' store the positins in an array
        ReDim Preserve wkns(i - startingLine + 1)
        ReDim Preserve updated(i - startingLine + 1)
        wkns(i - startingLine + 1) = wkn
        updated(i - startingLine + 1) = False
        
        If DateTime.Date <> Int(wsFunds.Cells(i, clmDate)) Then
            If favorites Then
                ' update only favorites
                If wsFunds.Cells(i, clmFavorites) <> "" Then
                    Call GetAriva_Fund(url, wkn, price, currency_, development, country, _
                    sector, benchmark, alpha, beta, sharpeRatio, volatility, trackingError, _
                    correlation, skewness, kurtosis, sortinoRatio, informationRatio, r2, _
                    treynorRatio)
                    
                    ' sometimes the performance do not exist at ariva.de
                    ' in this case it will be updated from finanzen.net:
                    If development(1) = "-" Or development(1) = "0" Or development(2) = "" Or development(3) = "" Then
                        url_Finanzen = wsFunds.Cells(i, clmUrl_finanzen)
                        Call GetFinanzen_Fund(url_Finanzen, wkn, development, currency_, country, _
                            benchmark)
                        wsFunds.Cells(i, clmUrl_finanzen) = url_Finanzen
                    End If
                    
                    Call Update_Fund(i, url, price, currency_, development, country, _
                    sector, benchmark, alpha, beta, sharpeRatio, volatility, trackingError, _
                    correlation, skewness, kurtosis, sortinoRatio, informationRatio, r2, _
                    treynorRatio)
                    
                    updated(i - startingLine + 1) = True
                End If
            Else
                ' update all funds
                Call GetAriva_Fund(url, wkn, price, currency_, development, country, _
                    sector, benchmark, alpha, beta, sharpeRatio, volatility, trackingError, _
                    correlation, skewness, kurtosis, sortinoRatio, informationRatio, r2, _
                    treynorRatio)
                    
                ' sometimes the performance do not exist at ariva.de
                ' in this case it will be updated from finanzen.net:
                If development(1) = "-" Or development(1) = "0" Or development(2) = "" Or development(3) = "" Then
                    url_Finanzen = wsFunds.Cells(i, clmUrl_finanzen)
                    Call GetFinanzen_Fund(url_Finanzen, wkn, development, currency_, country, _
                        benchmark)
                    wsFunds.Cells(i, clmUrl_finanzen) = url_Finanzen
                End If
                
                Call Update_Fund(i, url, price, currency_, development, country, _
                    sector, benchmark, alpha, beta, sharpeRatio, volatility, trackingError, _
                    correlation, skewness, kurtosis, sortinoRatio, informationRatio, r2, _
                    treynorRatio)
                    
                updated(i - startingLine + 1) = True
            End If
        End If
    
        i = i + 1
    Wend
    
    Call Selenium_ariva.CloseSeleniumDriver
    Call Selenium_Finanzen.CloseSeleniumDriver
    
    ' sorting (3m)
    wsFunds.Cells(1, clmSortingIndex) = 1
    Sorting_Funds
    
    i = startingLine
    While wsFunds.Cells(i, clmWKN) <> ""
        ' find the wkn in an array:
        j = 1
        Do Until j > UBound(wkns)
            If wkns(j) = wsFunds.Cells(i, clmWKN) Then
                If updated(j) Then
                    wsFunds.Cells(i, clmChangeOfPosition) = j - (i - startingLine + 1)
                End If
                
                Exit Do
            End If
            
            j = j + 1
        Loop
        
        i = i + 1
    Wend
    
    Application.decimalSeparator = decimalSeparator
End Sub

Function Update_Fund(lineNr%, url$, price$, currency_$, development$(), country$, sector$, _
    benchmark$, alpha$, beta$, sharpeRatio$, volatility$, trackingError$, correlation$, _
    skewness$, kurtosis$, sortinoRatio$, informationRatio$, r2$, treynorRatio$)
    Dim i%
    
    wsFunds.Cells(lineNr, clmCurrency) = currency_
    wsFunds.Cells(lineNr, clmCountry) = country
    wsFunds.Cells(lineNr, clmSector) = sector
    wsFunds.Cells(lineNr, clmBenchmark) = benchmark
    
    If alpha = "" Then
        wsFunds.Cells(lineNr, clmAlpha) = ""
    Else
        wsFunds.Cells(lineNr, clmAlpha) = CDbl(alpha)
    End If
    
    If beta = "" Then
        wsFunds.Cells(lineNr, clmBeta) = ""
    Else
        wsFunds.Cells(lineNr, clmBeta) = CDbl(beta)
    End If
    
    If sharpeRatio = "" Then
        wsFunds.Cells(lineNr, clmSharpeRatio) = ""
    Else
        wsFunds.Cells(lineNr, clmSharpeRatio) = CDbl(sharpeRatio)
    End If
    
    If volatility = "" Then
        wsFunds.Cells(lineNr, clmVolatility) = ""
    Else
        wsFunds.Cells(lineNr, clmVolatility) = CDbl(volatility)
    End If
    
    If trackingError = "" Then
        wsFunds.Cells(lineNr, clmTrackingError) = ""
    Else
        wsFunds.Cells(lineNr, clmTrackingError) = CDbl(trackingError)
    End If
    
    If correlation = "" Then
        wsFunds.Cells(lineNr, clmCorrelation) = ""
    Else
        wsFunds.Cells(lineNr, clmCorrelation) = CDbl(correlation)
    End If
    
    If skewness = "" Then
        wsFunds.Cells(lineNr, clmSkewness) = ""
    Else
        wsFunds.Cells(lineNr, clmSkewness) = CDbl(skewness)
    End If
    
    If kurtosis = "" Then
        wsFunds.Cells(lineNr, clmKurtosis) = ""
    Else
        wsFunds.Cells(lineNr, clmKurtosis) = CDbl(kurtosis)
    End If
    
    If sortinoRatio = "" Then
        wsFunds.Cells(lineNr, clmSortinoRatio) = ""
    Else
        wsFunds.Cells(lineNr, clmSortinoRatio) = CDbl(sortinoRatio)
    End If
    
    If informationRatio = "" Then
        wsFunds.Cells(lineNr, clmInformationRatio) = ""
    Else
        wsFunds.Cells(lineNr, clmInformationRatio) = CDbl(informationRatio)
    End If
    
    If r2 = "" Then
        wsFunds.Cells(lineNr, clmR2) = ""
    Else
        wsFunds.Cells(lineNr, clmR2) = CDbl(r2)
    End If
    
    If treynorRatio = "" Then
        wsFunds.Cells(lineNr, clmTreynorRatio) = ""
    Else
        wsFunds.Cells(lineNr, clmTreynorRatio) = CDbl(treynorRatio)
    End If
    
    wsFunds.Cells(lineNr, clmUrl_ariva) = url
    
    If price = "" Then
        wsFunds.Cells(lineNr, clmPrice) = ""
    Else
        wsFunds.Cells(lineNr, clmPrice) = CDbl(Replace(price, ".", ""))
    End If
    
    For i = 1 To 5
        If development(i) = "-" Or development(i) = "" Then
            wsFunds.Cells(lineNr, clm_3_months + i - 1) = ""
        Else
            development(i) = Replace(development(i), "%", "")
            wsFunds.Cells(lineNr, clm_3_months + i - 1) = CDbl(development(i)) / 100#
        End If
    Next i
    
    wsFunds.Cells(lineNr, clmDate) = DateTime.Now
End Function

Public Sub Sorting_Funds()
    Dim sortingRange As Range
    Dim nRows&, nColumns&, sortedColumn&
    Dim strRange As String
    Dim offset%
    
    Set wsFunds = Worksheets("Funds")
    
    nRows = 1
    While wsFunds.Cells(nRows + 2, 1) <> ""
        nRows = nRows + 1
    Wend
    nRows = nRows - 1
    
    offset = Application.WorksheetFunction.Match("Sorting", wsFunds.Range("A:A"), 0)
    
    sortedColumn = wsFunds.Cells(offset + wsFunds.Cells(1, clmSortingIndex), 2)
    
    nColumns = wsFunds.Cells(2, wsFunds.Columns.Count).End(xlToLeft).Column
    
    strRange = wsFunds.Range(wsFunds.Cells(3, 1), wsFunds.Cells(nRows + 2, nColumns)).Address(ReferenceStyle:=xlA1)
    
    wsFunds.Sort.SortFields.Clear
    
    Set sortingRange = _
        Range(wsFunds.Range(wsFunds.Cells(3, sortedColumn), wsFunds.Cells(nRows + 2, sortedColumn)).Address(ReferenceStyle:=xlA1))
    
    If wsFunds.Cells(offset + wsFunds.Cells(1, clmSortingIndex), 3) Then
        wsFunds.Sort.SortFields.Add Key:=sortingRange, _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    Else
        wsFunds.Sort.SortFields.Add Key:=sortingRange, _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    End If
    
    With wsFunds.Sort
        .SetRange Range(strRange)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
