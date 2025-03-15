
<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<%

	Site_Code = request("Site_Code")
	include_kbn = request("include_kbn")
	

	sql = "IF EXISTS( SELECT  * " & _
		"           FROM    Prj_Include_Chk " & _
		"           WHERE   Site_Code = '" & Site_Code & "') " & _
		"   UPDATE  Prj_Include_Chk " & _
		"   SET     Include_KBN = '" & include_kbn & "' " & _
		"   WHERE   Site_Code = '" & Site_Code & "' " & _
		"ELSE " & _
		"   INSERT Prj_Include_Chk(Site_Code , Include_KBN) VALUES('" & _
			Site_Code & "','" & include_kbn & "') "
			
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