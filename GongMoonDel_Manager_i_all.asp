<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="misop.inc" -->
<%

Str = Request("Str")
Page = Request("Page")


db_id = session("db_id")

c_date = year(Date)
type__  = Request("type__")

Strlist   = split(Str,   ";")



		   Set DbRec_Mi=Server.CreateObject("ADODB.Recordset")
		   DbRec_Mi.CursorType=1



			sql1 = "select db_id=c_userid, db_passwd=c_passwd, db_name=c_name, "
			sql1 = sql1 & " db_email=c_email,  db_level=c_level, db_longname=c_longname "
			sql1 = sql1 & "  from USER_TBL  "
			sql1 = sql1 & "  where c_level = 'S' "

			DbRec_Mi.Open sql1, DbCon

			r_count = DbRec_Mi.Recordcount

				If UBound(Strlist) > 0 then
					For i=0 To UBound(Strlist)-1

							    sqlstr = "IF EXISTS( SELECT  * " & _
							             "           FROM    OFFICE_MANAGER_i " & _
							             "           WHERE   seq = " & Strlist(i) & " AND s_id='" & db_id & "' ) " & _
							             "   UPDATE  OFFICE_MANAGER_i " & _
							             "   SET     Del_Chk='1' " & _
							             "   WHERE   seq = " & Strlist(i) & " AND s_id='" & db_id & "'" & _
							             " ELSE " & _
							             " INSERT OFFICE_MANAGER_i(seq,s_id,Del_Chk) values(" & Strlist(i) & ",'" & db_id & "','1') "

								Set Result = DbCon.execute(sqlstr)
								Set Result=Nothing


							  if r_count <> 0 then
								  DbRec_Mi.MoveLast

									FOR H = 1 TO r_count
									
										sqlstr = "IF EXISTS( SELECT  * " & _
												 "           FROM    OFFICE_MANAGER_i " & _
												 "           WHERE   seq = " & Strlist(i) & " AND s_id='" & DbRec_Mi("db_id") & "' ) " & _
												 "   UPDATE  OFFICE_MANAGER_i " & _
												 "   SET     Del_Chk='1' " & _
												 "   WHERE   seq = " & Strlist(i) & " AND s_id='" & DbRec_Mi("db_id") & "'" & _
												 " ELSE " & _
												 " INSERT OFFICE_MANAGER_i(seq,s_id,Del_Chk) values(" & Strlist(i) & ",'" & DbRec_Mi("db_id") & "','1') "

										Set Result = DbCon.execute(sqlstr)
										Set Result=Nothing										

										DbRec_Mi.MovePrevious
									NEXT

							  end if



					Next
				End If

				Set DbRec_Mi= Nothing


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
//alert(CPage);
	//window.location.href = "right_main_gongmoon.asp?Seq=" + pItemBoardID + "&Page=" + CPage;
	parent.frames[2].location = "ilban_gongmoon_manager.asp?Page=" + <%=Page%> + "&type__=<%=type__%>";
	//parent.frames[0].location = "right_main_gongmoon.asp?Page=" + <%=Page%>;
}
-->
</script>

</HEAD>
<!--BODY onload="PrintFrm()"-->
<BODY onload="PrintFrm()">



</BODY>
</HTML>