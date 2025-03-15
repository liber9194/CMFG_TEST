<%@ LANGUAGE="VBSCRIPT" %>
<!-- #include file="../../../../default_properties.asp" -->
<%
'file = Request("file")
'
'url="http://210.122.146.200/cug_data/" & file
'Response.Redirect url
%>

<%   Option Explicit
  

    Dim strFilename, strFilePath,strDirectory
    Dim s, fso, f, intFilelength ,name,nname,file

    file = Request("file")

   'strDirectory = Server.MapPath("/sfg/Net_test/FrontEnd/Home_myoffice_SubModule/Mis/WebIMGs")
   'strFilePath = strDirectory & "\" & file
	strDirectory = g_file_real_path & "/WebIMGs"
	strFilePath = strDirectory & "/" & file	
   
  

   
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

	if lcase(right(trim(file),3)) = "hwp" then
		Response.CharSet = "UDF-8"  
	else
	    Response.CharSet = "UTF-8"  
	end if

    Response.ContentType = "application/octet-stream"  

    Response.BinaryWrite s.Read  

    Response.Flush  

    s.Close  

    Set s = Nothing  

%>








<% 
'        Response.ContentType = "application/octect-stream name="& file

'        Response.AddHeader "Content-disposition", "attachment; filename="& file

        

'        Set objStream = Server.CreateObject("Adodb.Stream")

'        objStream.Open

'        objStream.LoadFromFile(dir & fn)

 '       buff = objStream.ReadText(-1)

 '       Response.BinaryWrite buff

        

  '      Set buff = Nothing

   '     Set objStream = Nothing

%>










<% 

'filename = Request("file") 

'filepath = "\down\"

'strPath = Server.Mappath("/") & filepath & Filename

'Response.AddHeader "Content-Disposition","attachment;filename=" & filename 

'Response.ContentType = "application/unknown" 

'Response.CacheControl = "private" 

'set objFS =Server.CreateObject("Scripting.FileSystemObject") 

'set objF = objFS.GetFile(strPath) 

'Response.AddHeader "Content-Length", objF.Size 

'set objF = nothing 

'set objFS = nothing 

'Set objStream = Server.CreateObject("ADODB.Stream")

'objStream.Open

'objStream.Type = 1

'objStream.LoadFromFile strPath

'download = objStream.Read

'Response.BinaryWrite download

'Set objstream = nothing

%> 
