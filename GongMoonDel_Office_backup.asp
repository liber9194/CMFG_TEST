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



				If UBound(Strlist) > 0 then
					For i=0 To UBound(Strlist)-1

								sqlstr = " Update office_tbl "
								sqlstr 	= sqlstr & " set office_del = 'T' "
								sqlstr 	= sqlstr & " WHERE o_seq = " & Strlist(i) & " "

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
	parent.frames[2].location = "right_main_gongmoon_insert.asp?Page=" + <%=Page%> + "&type__=<%=type__%>";
}
-->
</script>

</HEAD>
<!--BODY onload="PrintFrm()"-->
<BODY onload="PrintFrm()">



</BODY>
</HTML>