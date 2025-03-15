
<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<%

	Site_Code = request("Site_Code")
	BIGO = request("BIGO")
	

	sql = sql & " INSERT Article_01([Site_Code],[GUBUN_NUMBER],[GUBUN],[ITEM],[RENT_COMPANY],[REG_NUMBER],[COMPUTER_NUMBER],[MONITOR_NUMBER],[START_DATE],[END_DATE],COUNT_I,[BIGO]) VALUES("
	sql = sql & "'" & Site_Code & "',"
	sql = sql & "'" & GUBUN_NUMBER &"',"
	sql = sql & "'" & GUBUN & "',"
	sql = sql & "'" & ITEM & "',"
	sql = sql & "'" & RENT_COMPANY & "',"
	sql = sql & "'" & REG_NUMBER & "',"
	sql = sql & "'" & COMPUTER_NUMBER & "',"
	sql = sql & "'" & MONITOR_NUMBER & "',"
	sql = sql & "'" & START_DATE & "',"
	sql = sql & "'" & END_DATE & "',"
	sql = sql & "(select isnull(max(COUNT_I),0)+1 from Article_01 where GUBUN_NUMBER = '" & GUBUN_NUMBER &"' and Site_Code='" + Site_Code +"'),"
	sql = sql & "'" & BIGO & "')" & vbCrLf
	
	sql = " IF EXISTS (SELECT * FROM Article_Bigo WHERE Site_Code = '" & Site_Code & "') "
	sql = sql & " 	BEGIN "
	sql = sql & " 		UPDATE Article_Bigo "
	sql = sql & " 		SET Bigo = '" & Bigo & "' "
	sql = sql & "		WHERE Site_Code = '" & Site_Code & "' "
	sql = sql & " 	END "
	sql = sql & " ELSE "
	sql = sql & " 	BEGIN "
	sql = sql & " 		INSERT Article_Bigo(Site_Code, Bigo) "
	sql = sql & " 		VALUES('" & Site_Code & "', '" & Bigo & "') "
	sql = sql & " END "

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