strHDLocation = "c:\script"
ListFile = "C:\script\list.txt"

Set objFSOtexto = CreateObject("Scripting.FileSystemObject")
Set objFSO = Createobject("Scripting.FileSystemObject")
Set objADOStream = CreateObject("ADODB.Stream")
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

objADOStream.Open
objADOStream.Type = 1 'adTypeBinary

objADOStream.Write objXMLHTTP.ResponseBody
objADOStream.Position = 0    'Set the stream position to the start


objADOStream.SaveToFile strHDLocation & Year(now()) & right("0" & Month(now()),2) & right("0" & Day(now()),2) & right("0" & Hour(now()),2) & right("0" & Minute(now()),2) & right("0" & second(now()),2) & ".pdf"

objADOStream.Close


End if


Loop


Set objXMLHTTP = Nothing
Set objFSOtexto = Nothing
Set objADOStream = Nothing
objFile.Close
