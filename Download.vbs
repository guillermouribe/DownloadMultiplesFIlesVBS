strHDLocation = "c:\script\form"
ListFile = "C:\script\list.txt"

Set objFSOtexto = CreateObject("Scripting.FileSystemObject")
Set objFSO = Createobject("Scripting.FileSystemObject")
Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")


Set objFile = objFSOtexto.OpenTextFile(ListFile)
FileNumber = 0

Do Until objFile.AtEndOfStream
    URLFromList = objFile.ReadLine
    FileNumber = FileNumber + 1


' Fetch the file   
    objXMLHTTP.open "GET", URLFromList, false
    objXMLHTTP.send()

If objXMLHTTP.Status = 200 Then

Set objADOStream = CreateObject("ADODB.Stream")
objADOStream.Open
objADOStream.Type = 1 'adTypeBinary

objADOStream.Write objXMLHTTP.ResponseBody
objADOStream.Position = 0    'Set the stream position to the start



strHDLocation = strHDLocation & "_" & FileNumber & "_"
objADOStream.SaveToFile strHDLocation & Year(now()) & "-" & right("0" & Month(now()),2) & "-" & right("0" & Day(now()),2) & "_" & right("0" & Hour(now()),2) & right("0" & Minute(now()),2) & right("0" & second(now()),2) & ".pdf"

strHDLocation = "c:\script\form"
objADOStream.Close



End if



  

Loop


'objFSOtexto = Nothing
Set objXMLHTTP = Nothing
Set objFSOtexto = Nothing
'Set objADOStream = Nothing

'objFile.Close
