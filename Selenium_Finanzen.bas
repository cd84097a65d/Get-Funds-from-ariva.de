Attribute VB_Name = "Selenium_Finanzen"
Option Explicit

Const FinanzenNet = "https://www.finanzen.net/"
Const consentUUID = "45a9d072-f0f9-48bb-92e8-04e324cad5af"

Private SeleniumDriver_finanzen As ChromeDriver
Private seleniumStarted As Boolean

Dim keys As New Selenium.keys
Public cookiesAccepted As Boolean

Function GetFinanzen_Fund(url$, wkn$, development$(), currency_$, country$, benchmark$)
    Dim tmpWebElements As WebElements
    Dim tmpWebElement As WebElement
    Dim tmpString As String
    Dim tmpStrings() As String
    Dim i%
    
    If SeleniumDriver_finanzen Is Nothing Then
        ' create new selenium driver
        Set SeleniumDriver_finanzen = New ChromeDriver
        
        seleniumStarted = True
    End If
    
    ' if "finanzen_address" is not empty, then open it
    If url = "" Then
        ' otherwise, find fund according to WKN
        Call SeleniumDriver_finanzen.Get(FinanzenNet)
        Call Sleep(5000)
        
        AcceptCookies
        
        Set tmpWebElements = _
            SeleniumDriver_finanzen.FindElementsByXPath("//div[@id='search']/form/input")
        
        Call tmpWebElements(1).SendKeys(wkn, keys.Enter)
        
        ' go to "performance": replace
        ' https://www.finanzen.net/fonds/green-benefit-global-impact-fund-p-lu1136260384
        ' with
        ' https://www.finanzen.net/fonds/performance/green-benefit-global-impact-fund-p-lu1136260384
        url = Replace(SeleniumDriver_finanzen.url, "www.finanzen.net/fonds/", _
            "www.finanzen.net/fonds/performance/")
        Call SeleniumDriver_finanzen.Get(url)
    Else
        Call Sleep(5000)
        Call SeleniumDriver_finanzen.Get(url)
    
        AcceptCookies
    End If
    
    ' distinguish between fonds and ETFs:
    If InStr(1, url, "www.finanzen.net/etf/") > 0 Then
        tmpString = FindElementsByXPath(SeleniumDriver_finanzen, "//div[7]/h2", "development")
        If InStr(1, tmpString, "Anlageziel") > 0 Then
            i = 9
        Else
            i = 10
        End If
        
        currency_ = FindElementsByXPath(SeleniumDriver_finanzen, "//div/div/div[4]/div[2]/div", "Currency")
        country = FindElementsByXPath(SeleniumDriver_finanzen, "//div[20]/div/div[2]/div[2]/div[2]", "Country")
        benchmark = FindElementsByXPath(SeleniumDriver_finanzen, "//div[5]/div[2]/div", "Benchmark")
        
        development(1) = CheckValue_prc(FindElementsByXPath(SeleniumDriver_finanzen, "//div[" & i & "]/div/table/tbody/tr/td[2]", "3m"))
        development(2) = CheckValue_prc(FindElementsByXPath(SeleniumDriver_finanzen, "//div[" & i & "]/div/table/tbody/tr/td[3]", "6m"))
        development(3) = CheckValue_prc(FindElementsByXPath(SeleniumDriver_finanzen, "//div[" & i & "]/div/table/tbody/tr/td[4]", "1yr"))
        development(4) = CheckValue_prc(FindElementsByXPath(SeleniumDriver_finanzen, "//div[" & i & "]/div/table/tbody/tr/td[5]", "3yrs"))
        development(5) = CheckValue_prc(FindElementsByXPath(SeleniumDriver_finanzen, "//div[" & i & "]/div/table/tbody/tr/td[6]", "5yrs"))
        
        tmpString = tmpString
    Else
        tmpString = FindElementsByXPath(SeleniumDriver_finanzen, "//div[2]/table/tbody/tr", "development")
        tmpStrings = Split(tmpString, " ")
        tmpString = tmpString
        
        If UBound(tmpStrings) >= 3 Then
            development(1) = CheckValue(tmpStrings(3))
        Else
            development(1) = ""
        End If
        
        If UBound(tmpStrings) >= 4 Then
            development(2) = CheckValue(tmpStrings(4))
        Else
            development(2) = ""
        End If
        
        If UBound(tmpStrings) >= 6 Then
            development(3) = CheckValue(tmpStrings(6))
        Else
            development(3) = ""
        End If
        
        If UBound(tmpStrings) >= 8 Then
            development(4) = CheckValue(tmpStrings(8))
        Else
            development(4) = ""
        End If
        
        If UBound(tmpStrings) >= 10 Then
            development(5) = CheckValue(tmpStrings(10))
        Else
            development(5) = ""
        End If
    End If
    
    DoEvents
    Call Sleep(RandomRange(5000, 15000))
    DoEvents
End Function

' accept cookies
Sub AcceptCookies()
    Dim tmpWebElement As WebElement
    Dim cookies
    
    If Not cookiesAccepted Then
        Call SeleniumDriver_finanzen.Manage.AddCookie("consentUUID", _
            consentUUID, "www.finanzen.net")
        
        Call SeleniumDriver_finanzen.Get(SeleniumDriver_finanzen.url)
        
        cookiesAccepted = True
    End If
End Sub

Function CheckValue_prc$(inputValue$)
    If inputValue = "-" Or inputValue = "" Then
        CheckValue_prc = ""
    Else
        CheckValue_prc = Replace(Split(inputValue, "%")(0), " ", "")
    End If
End Function

Function CheckValue$(inputValue$)
    If inputValue = "-" Or inputValue = "" Then
        CheckValue = ""
    Else
        CheckValue = inputValue
    End If
End Function

Sub CloseSeleniumDriver()
    If Not (SeleniumDriver_finanzen Is Nothing) Then
        SeleniumDriver_finanzen.Quit
    End If
End Sub
