Sub DownloadFile()

Dim myURL As String
myURL = "https://cdn.discordapp.com/avatars/1053699929941491775/0d44a4860c30426cfe129251d256dd61.webp?size=100"

Dim WinHttpReq As Object
Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
WinHttpReq.Open "GET", myURL, False, "username", "password"
WinHttpReq.send

If WinHttpReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write WinHttpReq.responseBody
    oStream.SaveToFile "C:\Users\root\Desktop\Word\file.png", 2 ' 1 = no overwrite, 2 = overwrite
    oStream.Close
End If
