
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
	RENT_COMPANY = request("RENT_COMPANY")
	CAR_NUMBER = request("CAR_NUMBER")
	CAR_NAME = request("CAR_NAME")
	START_DATE = request("START_DATE")
	END_DATE = request("END_DATE")
	BIGO = request("BIGO")
	

	sql = sql & " INSERT Article_03([Site_Code],[GUBUN_NUMBER],[GUBUN],[ITEM],[RENT_COMPANY],[CAR_NUMBER],[CAR_NAME],[START_DATE],[END_DATE],COUNT_I,[BIGO])  VALUES("
	sql = sql & "'" & Site_Code & "',"
	sql = sql & "'" & GUBUN_NUMBER &"',"
	sql = sql & "'" & GUBUN & "',"
	sql = sql & "'" & ITEM & "',"
	sql = sql & "'" & RENT_COMPANY & "',"
	sql = sql & "'" & CAR_NUMBER & "',"
	sql = sql & "'" & CAR_NAME & "',"
	sql = sql & "'" & START_DATE & "',"
	sql = sql & "'" & END_DATE & "',"
	sql = sql & "(select isnull(max(COUNT_I),0)+1 from Article_03 where GUBUN_NUMBER = '" & GUBUN_NUMBER &"' and Site_Code='" + Site_Code +"'),"
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