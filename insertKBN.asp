
<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<%

	Site_Code = request("Site_Code")
	Del_KBN = request("Del_KBN")
	

	sql = "IF EXISTS( SELECT  * " & _
		"           FROM    Prj_Del_Chk " & _
		"           WHERE   Site_Code = '" & Site_Code & "') " & _
		"   UPDATE  Prj_Del_Chk " & _
		"   SET     Del_KBN = '" & Del_KBN & "' " & _
		"   WHERE   Site_Code = '" & Site_Code & "' " & _
		"ELSE " & _
		"   INSERT Prj_Del_Chk(Site_Code , Del_KBN) VALUES('" & _
			Site_Code & "','" & Del_KBN & "') "
			
	Set Result = DbCon.execute(sql)
	Set Result=Nothing
	DbCon.close
	m_sus = "ok"
		
%>
<HTML>
<HEAD>
<TITLE>Save</TITLE>
<script language="JavaScript">
var g_ExchangeVS = "<%=m_sus%>";
function fnMessage(){
	if (g_ExchangeVS == 'ok') {
		alert("저장되었습니다.");
		parent.goSearch();
	}
}
</script>

</HEAD>
<BODY onload="fnMessage()">
</BODY>
</HTML>