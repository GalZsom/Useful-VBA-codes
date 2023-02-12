Attribute VB_Name = "HTML_EUR_Query"
Public LatestMNBDate As String

Sub GetCurrentEURValue() 'Button subroutine
    Call EurValMsgBox
End Sub


Sub EurValMsgBox() 'The message
    MsgBox "Latest EUR-HUF : " & GetEurHuf & vbNewLine & LatestMNBDate & vbNewLine & "Source: https://www.mnb.hu/arfolyamok", vbInformation
End Sub

Function GetEurHuf()
    Dim html As Object
    Dim http As Object
    Dim htmlTable As Object
    Dim htmlRow As Object
    Dim htmlCell As Object
    Dim url As String
    
    Set html = CreateObject("htmlfile")
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://www.mnb.hu/arfolyamok"
    
    On Error GoTo ErrHandler 'Error handling
        With http
            .Open "GET", url, False
            .send
            If .Status = 200 Then html.body.innerHTML = .responseText
        End With
    On Error GoTo 0
    
    Set htmlTable = html.getElementsByTagName("table")(0)

    For Each htmlRow In htmlTable.getElementsByTagName("tr") 'Iterating through all the HTML cells
        For Each htmlCell In htmlRow.Children
            If htmlCell.innerText = CStr("EUR") Then
                GetEurHuf = htmlRow.getElementsByTagName("td")(3).innerText 'Returns current exchange rate
                Exit For
            End If
        Next htmlCell
    Next htmlRow
    
    LatestMNBDate = html.getElementsByTagName("caption")(0).innerText 'Returns the latest date of update
    
    If GetEurHuf = "" Then GetEurHuf = "HIBA, ilyen érték nem található" 'EH

Exit Function

ErrHandler:
    If Not http Is Nothing Then Set http = Nothing
    GetEurHuf = "Could not connect to " & url & "."
End Function
