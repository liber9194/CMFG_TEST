<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="misop.inc" -->
<!-- #include file="../../../../default_properties.asp" -->
<%

Str = Request("Str")
Page = Request("Page")


db_id = session("db_id")

c_date = year(Date)
type__  = Request("type__")



Strlist   = split(Str,   ";")




				If UBound(Strlist) > 0 then
					For i=0 To UBound(Strlist)-1


							Set DbRec=Server.CreateObject("ADODB.Recordset")
							DbRec.CursorType=1

							str = " SELECT o_savehtml  "
							str = str & " FROM save_html_i  "
							str = str & " WHERE o_seq =" & Strlist(i) & " "

							DbRec.Open str, DbCon


							if DbRec.Recordcount <> 0 then
								DbRec.MoveLast
								
								savehtml = DbRec("o_savehtml")
								

										 
								'filePath = server.MapPath("../Mis/WebIMGs/" & savehtml & "/") 	
								'del_folder = server.mappath("../Mis/WebIMGs/" & savehtml )
								filePath = g_file_real_path & "/WebIMGs/" & savehtml & "/"
								del_folder = g_file_real_path & "/WebIMGs/" & savehtml
								
								Set FSO = CreateObject("Scripting.FileSystemObject") 
									 
									If FSO.FolderExists(filePath) Then

										Set objFdr = FSO.GetFolder(filePath)

											For Each objFdr In objFdr.Files 

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

							Set DbRec = Nothing





							Set DbRec=Server.CreateObject("ADODB.Recordset")
							DbRec.CursorType=1

							str = " SELECT o_savehtml  "
							str = str & " FROM user_html  "
							str = str & " WHERE o_seq =" & Strlist(i) & " "

							DbRec.Open str, DbCon

							if DbRec.Recordcount <> 0 then
								DbRec.MoveLast



								for ii = 1 to DbRec.Recordcount
								
										savehtml = DbRec("o_savehtml")
																						 
										'filePath = server.MapPath("../Mis/WebIMGs/" & savehtml & "/") 	
										'del_folder = server.mappath("../Mis/WebIMGs/" & savehtml )
										filePath = g_file_real_path & "/WebIMGs/" & savehtml & "/"
										del_folder = g_file_real_path & "/WebIMGs/" & savehtml
										
										Set FSO = CreateObject("Scripting.FileSystemObject") 
											 
											If FSO.FolderExists(filePath) Then

												Set objFdr = FSO.GetFolder(filePath)

													For Each objFdr In objFdr.Files 

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
									DbRec.MovePrevious
								next

							end if

							Set DbRec = Nothing



							sqlstr = ""
							sqlstr 	= sqlstr & " delete office_tbl_i where o_seq = " & Strlist(i) & " "
							sqlstr 	= sqlstr & " delete RECEIVE_TBL_i where o_seq = " & Strlist(i) & " "
							sqlstr 	= sqlstr & " delete save_html_i where o_seq = " & Strlist(i) & " "
							sqlstr 	= sqlstr & " delete save_file_i where o_seq = " & Strlist(i) & " "
							sqlstr 	= sqlstr & " delete user_html where o_seq = " & Strlist(i) & " "
							sqlstr 	= sqlstr & " delete user_file where o_seq = " & Strlist(i) & " "


								Set Result = DbCon.execute(sqlstr)
								Set Result=Nothing


					Next
				End If



%>

<HTML>
<HEAD>
<TITLE>Del</TITLE>


<script language="JavaScript">
<!--
function PrintFrm(){
	var CPage = "<%=Page%>"
	parent.frames[2].location = "ilban_gongmoon_write.asp?Page=" + <%=Page%> + "&type__=<%=type__%>";
}
-->
</script>

</HEAD>
<!--BODY onload="PrintFrm()"-->
<BODY onload="PrintFrm()">



</BODY>
</HTML>