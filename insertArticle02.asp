
<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<%

	Site_Code = request("Site_Code")
	GUBUN_NUMBER = request("GUBUN_NUMBER")
	GUBUN = request("GUBUN")
	ITEM = request("ITEM")
	SU_NUM = request("SU_NUM")
	GUIP = request("GUIP")
	GUIP_DATE = request("GUIP_DATE")
	SINGU = request("SINGU")
	BIGO = request("BIGO")
	

	sql = sql & " INSERT Article_02([Site_Code],[GUBUN_NUMBER],[GUBUN],[ITEM],[SU_NUM],[GUIP],[GUIP_DATE],[SINGU],COUNT_I,[BIGO])  VALUES("
	sql = sql & "'" & Site_Code & "',"
	sql = sql & "'" & GUBUN_NUMBER &"',"
	sql = sql & "'" & GUBUN & "',"
	sql = sql & "'" & ITEM & "',"
	sql = sql & "'" & SU_NUM & "',"
	sql = sql & "'" & GUIP & "',"
	sql = sql & "'" & GUIP_DATE & "',"
	sql = sql & "'" & SINGU & "',"
	sql = sql & "(select isnull(max(COUNT_I),0)+1 from Article_02 where GUBUN_NUMBER = '" & GUBUN_NUMBER &"' and Site_Code='" + Site_Code +"'),"
	sql = sql & "'" & BIGO & "')" & vbCrLf

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
		alert("추가되었습니다.");
		parent.goSearch();
	}
}
</script>

</HEAD>
<BODY onload="fnMessage()">
</BODY>
</HTML>