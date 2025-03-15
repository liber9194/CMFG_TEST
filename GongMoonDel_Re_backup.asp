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


				If UBound(Strlist) > 0 then

					For i=0 To UBound(Strlist)-1


							Set DbRec1=Server.CreateObject("ADODB.Recordset")
							DbRec1.CursorType=1

							str = " SELECT d_gubun_str  "
							str = str & " FROM REQUEST_BOARD_TBL  "
							str = str & " WHERE index_ =" & Strlist(i) & " "

							DbRec1.Open str, DbCon


							if DbRec1.Recordcount <> 0 then
								DbRec1.MoveLast
								
								d_gubun_str = DbRec1("d_gubun_str")
								

									Set DbRec2=Server.CreateObject("ADODB.Recordset")
									DbRec2.CursorType=1

									str = " SELECT d_gubun_str  "
									str = str & " FROM REQUEST_BOARD_TBL  "
									str = str & " WHERE d_gubun_str like '" & Strlist(i) & "%' and d_gubun_str <> '" & Strlist(i) & "' "

									DbRec2.Open str, DbCon
								
									'response.write str
									if DbRec2.Recordcount <> 0 then
										DbRec2.MoveLast
										%>
											<script language="JavaScript">
												alert("답변이 있는 글은 삭제 할수 없습니다....");
											</script>
										<%
										exit for

										chk_is = true
									else
										chk_is = false
									end if
							else
								
								d_gubun_str = ""
								chk_is = true
							end if
						

							if chk_is = false then

										Set DbRec=Server.CreateObject("ADODB.Recordset")
										DbRec.CursorType=1

										str = " SELECT d_savehtml  "
										str = str & " FROM REQUEST_savehtml  "
										str = str & " WHERE d_index =" & Strlist(i) & " "

										DbRec.Open str, DbCon


										if DbRec.Recordcount <> 0 then
											DbRec.MoveLast
											
											savehtml = DbRec("d_savehtml")
											
										else
											
											savehtml = ""
											
										end if

										sqlstr = " delete REQUEST_BOARD_TBL WHERE index_ =" & Strlist(i) & " "
										sqlstr 	= sqlstr & " delete REQUEST_savehtml where d_index =" & Strlist(i) & " "
										sqlstr 	= sqlstr & " delete REQUEST_savefile where d_index =" & Strlist(i) & " "

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
							end if





					Next
				End If



%>

<HTML>
<HEAD>
<TITLE>Del</TITLE>


<script language="JavaScript">
<!--
function PrintFrm(){
//	window.close();
//	opener.window.location.href="Right_Main_GongMoon.asp";
//	parent.right.location.href="Rail_List.asp";
//}
	//var firstWin = window.parent.opener;
	//firstWin.location = "right_main_gongmoon.asp?Seq=" + pItemBoardID + "&Page=" + <%=Page%>;
	//alert("t");
	//window.close();






	var CPage = "<%=Page%>"
	var Ccode = "<%=code%>"
	var CQstr = "<%=Qstr%>"
//alert(CPage);
//alert(Ccode);
	//parent.frames[1].location = "Right_JaRyo_Upmubogo.asp?Page=" + CPage + "&code=" + Ccode;
	
	
	
	
	//if ( "<%=tty%>" == "1" ) {
		
		
		
		
		parent.frames[2].location = "Right_JaRyo_Re.asp?Page=" + CPage + "&code=" + Ccode + "&Qstr=" + CQstr;

	//} else {
	//	parent.frames[2].location = "Right_JaRyo_Upmubogo.asp?Page=" + CPage + "&code=" + Ccode + "&Qstr=" + CQstr;
	//}








//	parent.frames[2].location = "Right_JaRyo_Upmubogo.asp?Page=" + CPage + "&code=" + Ccode + "&Qstr=" + CQstr;

	//parent.frames[2].location = "right_JaRyo_Upmubogo.asp?code=" + Pcode + "&Qstr=" + Pstr;
}
-->
</script>

</HEAD>
<!--BODY onload="PrintFrm()"-->
<BODY onload="PrintFrm()">



</BODY>
</HTML>