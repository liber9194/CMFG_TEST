<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="misop.inc" -->
<!-- #include file="../../../../default_properties.asp" -->
<%

Str = Request("Str")
Page = Request("Page")
code = Request("code")
Qstr = Request("Qstr")
tty = Request("tty")


db_id = session("db_id")

c_date = year(Date)




Strlist   = split(Str,   ";")
							
							
							'response.write Str
							'response.write "<br>"

'response.write code

'response.write "<br>"
				If UBound(Strlist) > 0 then

					For i=0 To UBound(Strlist)-1


							Set DbRec=Server.CreateObject("ADODB.Recordset")
							DbRec.CursorType=1

							str = " SELECT d_savehtml  "
							str = str & " FROM dk_savehtml  "
							str = str & " WHERE (dk_code = '" & code & "')  and  d_seq =" & Strlist(i) & " "

							DbRec.Open str, DbCon


							if DbRec.Recordcount <> 0 then
								DbRec.MoveLast
								
								savehtml = DbRec("d_savehtml")
								
							else
								
								savehtml = ""
								
							end if



							sqlstr = " delete PDS_TBL WHERE dk_code = '" & code & "' and  d_seq =" & Strlist(i) & " "
							sqlstr 	= sqlstr & " delete dk_savehtml where dk_code = '" & code & "' and  d_seq =" & Strlist(i) & " "
							sqlstr 	= sqlstr & " delete dk_savefile where dk_code = '" & code & "' and  d_seq =" & Strlist(i) & " "

							Set Result = DbCon.execute(sqlstr)
							Set Result=Nothing
							
							
							'response.write sqlstr
							'response.write "<br>"


							'response.write savehtml

							if savehtml <> "" then
							
									Dim FSO, filePath, objFdr
										 
									'filePath = server.MapPath("../Mis/WebIMGs/" & savehtml & "/") 	
									'del_folder = server.mappath("../Mis/WebIMGs/" & savehtml )
									filePath = g_file_real_path & "/WebIMGs/" & savehtml & "/" 
									del_folder = g_file_real_path & "/WebIMGs/" & savehtml

									'response.write "<br>filePath: " & filePath
									'response.write "<br>del_folder: " & del_folder

									Set FSO = CreateObject("Scripting.FileSystemObject") 
										 
									If FSO.FolderExists(filePath) Then

										'response.write "<br>폴더 존재"
										
										Set objFdr = FSO.GetFolder(filePath)

											For Each objFdr In objFdr.Files 
												'response.write "<br>파일 존재 : " & objFdr
												'fnm = ""

												'fnm = FSO.GetFileName(objFdr)
												FSO.DeleteFile(objFdr)
												
												'response.write fnm
												'response.write "<br>"
												'FSO.DeleteFile(fnm)

											Next
										
										FSO.deletefolder(del_folder)


										Set objFdr = Nothing
	
									End If
										 
									Set FSO = Nothing

							end if






					Next
				End If



%>

<HTML>
<HEAD>
<TITLE>Del</TITLE>


<script language="JavaScript">
function PrintFrm(){

	var CPage = "<%=Page%>"
	var Ccode = "<%=code%>"
	var CQstr = "<%=Qstr%>"

	if ( "<%=tty%>" == "1" ) {
		parent.frames[2].location = "Right_JaRyo_ver_up.asp?Page=" + CPage + "&code=" + Ccode + "&Qstr=" + CQstr;
	} else {
		parent.frames[2].location = "Right_JaRyo_Upmubogo.asp?Page=" + CPage + "&code=" + Ccode + "&Qstr=" + CQstr;
	}
}
</script>

</HEAD>

<BODY onload="PrintFrm()">



</BODY>
</HTML>