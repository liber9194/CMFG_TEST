<!-- #include file="../../../../default_properties.asp" -->
<%   
  

    Dim strFilename, strFilePath,strDirectory
    Dim s, fso, f, intFilelength ,name,nname,file

    file = Request("file_")
	dir1 = Request("dir_")
   'strDirectory = Server.MapPath("/sfg/Net_test/FrontEnd/Home_myoffice_SubModule/Mis/WebWrite_images")
   'strFilePath = strDirectory & "\" & dir1 & "\" &  file
   strDirectory = g_file_real_path & "/WebIMGs"
   strFilePath = strDirectory & "/" dir1 & "/"  & file
		
   
Response.write strFilePath

dim d
  d = 1

   if d =1 then
    Response.Buffer = True  

    Response.Clear  

    Set s = Server.CreateObject("ADODB.Stream")  

    s.Open  

    s.Type = 1  

    Set fso = Server.CreateObject("Scripting.FileSystemObject")  

    Set f = fso.GetFile(strFilePath)

    intFilelength = f.size  

    s.LoadFromFile(strFilePath) 

    

    Response.ContentType = "application/octect-stream name=" & file
    Response.AddHeader "Content-Disposition", "attachment; filename=" & file
    Response.AddHeader "Content-Length", intFilelength  
    Response.CharSet = "UTF-8"  
    Response.ContentType = "application/octet-stream"  
    Response.BinaryWrite s.Read  



'nDpSize = 1048576
'nDpSize =  1024768
'nLoopCnt = s.Size / nDpSize

 

'For i = 1 To nLoopCnt
'   Response.BinaryWrite s.Read(nDpSize)
'Next

 

'루프에서 처리 안된 나머지 값.
'If (s.Size Mod nDpSize) <> 0 Then
'   Response.BinaryWrite oStream.Read(s.Size Mod nDpSize)
'End If



    Response.Flush  
    s.Close  



	'if lcase(right(trim(file),3)) = "hwp" then
	'	Response.ContentType = "application/unknown name=" & file
	'	Response.AddHeader "Content-Disposition", "attachment; filename=" & file
	'	Response.AddHeader "Content-Length", intFilelength  
	'	Response.CharSet = "UDF-8"  
	'	Response.ContentType = "application/unknown"  
	'	Response.BinaryWrite s.Read  
	'	Response.Flush  
	'	s.Close  
	'else
	'	Response.ContentType = "application/octect-stream name=" & file
	'	Response.AddHeader "Content-Disposition", "attachment; filename=" & file
	'	Response.AddHeader "Content-Length", intFilelength  
	'	Response.CharSet = "UTF-8"  
	'	Response.ContentType = "application/octet-stream"  
	'	Response.BinaryWrite s.Read  
	'	Response.Flush  
	'	s.Close  
	'end if



    Set s = Nothing  


end if

%>
