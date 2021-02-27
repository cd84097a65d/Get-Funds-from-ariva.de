Attribute VB_Name = "Common"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Function RandomRange(lowerBound As Long, upperBound As Long) As Long
    Randomize
    RandomRange = CLng((upperBound - lowerBound + 1&) * Rnd + lowerBound)
End Function

Function FindElementsByXPath$(driver As ChromeDriver, xpath$, lineToFind)
    Dim nIterations%
    Dim tmpWebElements As WebElements
    
    FindElementsByXPath$ = ""
    Set tmpWebElements = driver.FindElementsByXPath(xpath)
    
    ' wait 10 iterations if the value is still not shown
    Do While tmpWebElements.Count = 0
        Set tmpWebElements = driver.FindElementsByXPath(xpath)
        nIterations = nIterations + 1

        If nIterations >= 10 Then
            Exit Function
        End If
    Loop

    FindElementsByXPath = tmpWebElements(1).Text
End Function

' only works for euro and dollar
Function ConvertCurrencySymbolToCurrencyName(inputString As String) As String
    Select Case inputString
       Case "Euro"
          ConvertCurrencySymbolToCurrencyName = "EUR"
       Case "€"
          ConvertCurrencySymbolToCurrencyName = "EUR"
       Case "US Dollar"
            ConvertCurrencySymbolToCurrencyName = "USD"
        Case "$"
            ConvertCurrencySymbolToCurrencyName = "USD"
       Case Else
          ConvertCurrencySymbolToCurrencyName = inputString
    End Select
End Function
