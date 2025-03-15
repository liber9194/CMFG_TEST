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
function PrintFrm(){
	var CPage = "<%=Page%>";
	parent.frames[2].location = "right_main_gongmoon_insert_CM.asp?Page=" + <%=Page%> + "&type__=<%=type__%>";
}
</script>

</HEAD>
<BODY onload="PrintFrm()">



</BODY>
</HTML>