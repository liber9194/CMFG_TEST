<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<!-- #include file="../../../../default_properties.asp" -->
<%	
	Site_Code = request("Site_Code")
	add_order_office_code = request("add_order_office_code")
	
	sql = sql & " insert prj_order_office_tbl_gamri "
	sql = sql & " values('" & Site_Code & "', " & add_order_office_code & ") "
	
	' for i = 1 to Request.Form("chk_tGita").count		
		' 'response.write Request.Form("chk_tGita")(i) & "<BR>"
		' sql = sql & " Delete From Article_03 where key_field = '" & Request.Form("chk_tGita")(i) & "' and site_code = '" & Site_Code & "' and GUBUN_NUMBER = '" & GUBUN_NUMBER & "' "
	' next	
	
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
		alert("추가되었습니다!");
		parent.goSearch();
	}
}
</script>

</HEAD>
<BODY onload="fnMessage()">
</BODY>
</HTML>	