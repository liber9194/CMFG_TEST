
<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<%

	db_id = session("db_id")
	seq = request("seq")

	IF db_id <> "" AND  seq <> "" Then 

			sql = sql & " IF EXISTS( SELECT  * " & _
					 "           FROM    OFFICE_MANAGER " & _
					 "           WHERE   seq = " & seq & " AND s_id='" & db_id & "' ) " & _
					 "   UPDATE  OFFICE_MANAGER " & _
					 "   SET     Del_Chk='1' " & _
					 "   WHERE   seq = " & seq & " AND s_id='" & db_id & "'" & _
					 " ELSE " & _
					 " INSERT OFFICE_MANAGER(seq,s_id,Del_Chk) values(" & seq & ",'" & db_id & "','1') "									
			'response.write sql
			'response.end
			Set Result = DbCon.execute(sql)
			Set Result=Nothing
			DbCon.close
			m_sus = "ok"
			
	Else
		m_sus = "not"	
	End IF
		
%>
<HTML>
<HEAD>
<TITLE>Save</TITLE>
<script language="JavaScript">
var g_ExchangeVS = "<%=m_sus%>";
function fnMessage(){
	if (g_ExchangeVS == 'not') {
		alert("삭제할 공문이 없습니다. 관리자에게 문의하십시오.");
	}	
	if (g_ExchangeVS == 'ok') {
		alert("삭제되었습니다!");
		parent.goSearch();
	}
}
</script>

</HEAD>
<BODY onload="fnMessage()">
</BODY>
</HTML>