Attribute VB_Name = "Selenium_ariva"
' original page is:
' https://github.com/cd84097a65d/Get-Funds-from-ariva.de
' you are free to use/modify/sell this table as you wish

Option Explicit

Const ArivaDe = "https://ariva.de/"
Const consentUUID = "e47e46c6-5b81-46cf-ace7-be5302056448"

Private SeleniumDriver_ariva As ChromeDriver
Private seleniumStarted As Boolean

Dim keys As New Selenium.keys
Public cookiesAccepted As Boolean

Function GetAriva_Fund(url$, wkn$, price$, currency_$, development$(), country$, sector$, _
    benchmark$, alpha$, beta$, sharpeRatio$, volatility$, trackingError$, correlation$, _
    skewness$, kurtosis$, sortinoRatio$, informationRatio$, r2$, treynorRatio$)
    
    Dim tmpWebElements As WebElements
    Dim tmpWebElement As WebElement
    Dim tmpString As String
    Dim tmpStrings() As String
    Dim i%
    
    price = "": currency_ = "": alpha = "": sharpeRatio = "": beta = "": country = ""
    benchmark = "": sector = "": volatility = "": trackingError = "": correlation = ""
    skewness = "": kurtosis = "": sortinoRatio = "": informationRatio = "": r2 = ""
    treynorRatio = ""
    development(1) = "": development(2) = "": development(3) = "": development(4) = ""
    development(5) = ""
    
    If SeleniumDriver_ariva Is Nothing Then
        ' create new selenium driver
        Set SeleniumDriver_ariva = New ChromeDriver
        
        seleniumStarted = True
    End If
    
    ' if "url" is empty, then search for fund according to WKN and update url
    If url = "" Then
        Call SeleniumDriver_ariva.Get(ArivaDe)
        Call Sleep(5000)
        
        AcceptCookies
        
        Set tmpWebElements = _
            SeleniumDriver_ariva.FindElementsByXPath("//input")
        
        Call tmpWebElements(1).SendKeys(wkn, keys.Enter)
        
        url = SeleniumDriver_ariva.url
    Else
        Call SeleniumDriver_ariva.Get(url)
        Call Sleep(5000)
    
        AcceptCookies
    End If
    
    development(1) = FindElementsByXPath(SeleniumDriver_ariva, "//tr[3]/td[2]", "3m")
    development(2) = FindElementsByXPath(SeleniumDriver_ariva, "//tr[4]/td[2]", "6m")
    development(3) = FindElementsByXPath(SeleniumDriver_ariva, "//tr[2]/td[4]/span", "1yr")
    development(4) = FindElementsByXPath(SeleniumDriver_ariva, "//tr[3]/td[4]", "3yrs")
    development(5) = FindElementsByXPath(SeleniumDriver_ariva, "//tr[4]/td[4]", "5yrs")
    price = FindElementsByXPath(SeleniumDriver_ariva, "//td/span/span", "price")
    
    ' some funds have different content of the page, then use finanzen.net to get them
    If development(2) = "" Or development(3) = "" Then
        Exit Function
    End If
    
    currency_ = ConvertCurrencySymbolToCurrencyName(GetElementFromTheList("//div[2]/div[2]/div/table/tbody/tr[", "]/td", "]/td[2]", "Fondswährung"))
    
    ' alpha
    alpha = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "Alpha")
    
    ' Sharpe ratio
    sharpeRatio = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "Sharpe-Ratio 1 Jahr")
    
    ' beta
    beta = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "Beta")
    
    ' country
    country = GetElementFromTheList("//div[4]/div[2]/div/table/tbody/tr[", "]/td", "]/td[2]", "Aufgelegt in")
    
    ' benchmark
    benchmark = GetElementFromTheList("//div[4]/div[2]/div/table/tbody/tr[", "]/td", "]/td[2]", "Benchmark")
    
    ' sector
    sector = GetElementFromTheList("//div[4]/div[2]/div/table/tbody/tr[", "]/td", "]/td[2]", "Kategorie")
    
    ' volatility
    volatility = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "Volatilität 1 Jahr")
    
    ' tracking error
    trackingError = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "Tracking Error")
    
    ' correlation
    correlation = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "Korrelation")
    
    ' skewness
    skewness = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "Schiefe")
    
    ' kurtosis
    kurtosis = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "Kurtosis")
    
    ' Sortino ratio
    sortinoRatio = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "Sortino Ratio")
    
    ' Information ratio
    informationRatio = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "Information Ratio")
    
    ' R2
    r2 = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "R-Squared")
    
    ' Treynor Ratio
    treynorRatio = GetElementFromTheList("//div[4]/div/div/table/tbody/tr[", "]/td", "]/td[2]", "Treynor Ratio")
    
    DoEvents
    Call Sleep(RandomRange(5000, 15000))
    DoEvents
End Function

' since the list of categories can have varying length, we have to make a loop
' over the whole list to find the category we need
Function GetElementFromTheList$(leftXpath$, rightXpathCategory$, rightXpathValue$, category$)
    Dim i%
    
    GetElementFromTheList = ""
    For i = 1 To 20
        If FindElementsByXPath(SeleniumDriver_ariva, _
            leftXpath & CStr(i) & rightXpathCategory, category) = category Then
            GetElementFromTheList = FindElementsByXPath(SeleniumDriver_ariva, _
                leftXpath & CStr(i) & rightXpathValue, category)
            Exit For
        End If
    Next i
End Function

Sub CloseSeleniumDriver()
    If Not (SeleniumDriver_ariva Is Nothing) Then
        SeleniumDriver_ariva.Close
        SeleniumDriver_ariva.Quit
    End If
End Sub

Sub AcceptCookies()
    Dim tmpWebElement As WebElement
    Dim cookies
    
    If Not cookiesAccepted Then
        Call SeleniumDriver_ariva.Manage.AddCookie("consentUUID", _
            consentUUID, ".ariva.de")
        
        Call SeleniumDriver_ariva.Get(SeleniumDriver_ariva.url)
        
        cookiesAccepted = True
    End If
End Sub

