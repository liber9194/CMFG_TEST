<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<!-- #include file="../../../../default_properties.asp" -->
<%	
	Site_Code = request("Site_Code")
	del_order_office_code = request("del_order_office_code")
	
	'현장게시현황 상주 기술자 소속 업데이트
	sql = " update Prj_Gamri_Tuip "	
	sql = sql & " set tuip_sosok_cd = null "
	sql = sql & " where Site_Code = '" & Site_Code & "' and tuip_sosok_cd = " & del_order_office_code
	
	'현장게시현황 비상주 기술자 소속 업데이트
	sql = sql & " update Prj_Gamri_Tuip_bi "	
	sql = sql & " set tuip_sosok_bi_cd = null "
	sql = sql & " where Site_Code = '" & Site_Code & "' and tuip_sosok_bi_cd = " & del_order_office_code
	
	sql = sql & " delete prj_order_office_tbl_gamri "
	sql = sql & " where Site_Code = '" & Site_Code & "' and order_office_code = " & del_order_office_code
	
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
		alert("삭제되었습니다.");
		parent.goSearch();
	}
}
</script>

</HEAD>
<BODY onload="fnMessage()">
</BODY>
</HTML>	