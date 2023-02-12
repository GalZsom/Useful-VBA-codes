Attribute VB_Name = "Office_Language_Detection"

Sub DetermineOfficeLanguage()
    Call GetXlLang

End Sub


Sub GetXlLang()
    Dim lngCode As Long
    lngCode = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    MsgBox "Code is: " & lngCode & vbNewLine & GetLocale(lngCode)
End Sub


Function GetLocale(ByVal lngCode) As String
    Dim html As Object
    Dim http As Object
    Dim htmlTable As Object
    Dim htmlRow As Object
    Dim htmlCell As Object
    Dim url As String
    
    Set html = CreateObject("htmlfile")
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://www.science.co.il/language/Locale-codes.php"
    
    On Error GoTo ErrHandler
        With http
            .Open "GET", url, False
            .send
            If .Status = 200 Then html.body.innerHTML = .responseText
        End With
    On Error GoTo 0
    
    Set htmlTable = html.getElementsByTagName("table")(0)

    For Each htmlRow In htmlTable.getElementsByTagName("tr")
        For Each htmlCell In htmlRow.Children
            If htmlCell.innerText = CStr(lngCode) Then
                GetLocale = htmlRow.getElementsByTagName("td")(0).innerText
                Exit For
            End If
        Next htmlCell
    Next htmlRow
    
    If GetLocale = "" Then GetLocale = "Value Not Found From " & url

Exit Function

ErrHandler:
    If Not http Is Nothing Then Set http = Nothing
    GetEurHuf = "Could not connect to " & url & "."
End Function
