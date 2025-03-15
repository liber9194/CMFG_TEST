
<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<%

	I_UserID = request("I_UserID")
	I_DongJel = request("I_DongJel")
	

	sql = "IF EXISTS( SELECT  * " & _
		"           FROM    Prj_DongJel " & _
		"           WHERE   UserID = '" & I_UserID & "') " & _
		"   UPDATE  Prj_DongJel " & _
		"   SET     DongJel = '" & I_DongJel & "' " & _
		"   WHERE   UserID = '" & I_UserID & "' " & _
		"ELSE " & _
		"   INSERT Prj_DongJel(UserID , DongJel) VALUES('" & I_UserID & "','" & I_DongJel & "') "
			
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