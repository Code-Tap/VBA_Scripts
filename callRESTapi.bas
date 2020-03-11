Sub apiTest()
    Dim oRequest As Object
    Set oRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    oRequest.Open "GET", "https://swapi.co/api/people/1/", False
    oRequest.SetRequestHeader "X-Auth-Token", "replace this with api token"
    oRequest.send
    Cells(1, 1) = oRequest.responseText
End Sub
