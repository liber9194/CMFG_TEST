
<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<%

	i_userid = request("i_userid")
	site_code = request("site_code")
	

	IF i_userid <> "" AND  site_code <> "" Then 
	    '중복체크
		sql = "  select count(*) as cnt from PM_manager where UserId = '" & i_userid & "' and site_code='" & site_code & "'"
		Set rs=Server.CreateObject("ADODB.Recordset")
		rs.CursorType=1 ' CursorTyp : 0 : adOpenForwardOnly (기본값) 	1 : adOpenKeyset 	2 : adOpenDynamic 	3 : adOpenStatic 
		cnt =0 
		rs.Open sql, DbCon
		If rs.Recordcount <> 0 Then 
			cnt = rs("cnt")
		End if
		
		'등록된 게 없는 경우 
		if cnt = 0 then
            sql = " INSERT INTO PM_Manager(Site_Code,UserID) VALUES ('" & site_code & "','" & i_userid & "'); "
            sql = sql & " INSERT INTO PM_Manager_His(Site_Code,UserID,IN_DATE,OUT_DATE,CNT_) VALUES ('" & site_code & "','" & i_userid & "',convert(varchar(10),getdate(),120),'',(SELECT  isnull(MAX(CNT_),0)+1 AS Cnt__ FROM  PM_Manager_His   WHERE   Site_Code = '" & site_code & "')); "
			
			'response.write sql
			'response.end
			Set Result = DbCon.execute(sql)
			Set Result=Nothing
			DbCon.close
			m_sus = "ok"
		else
			m_sus = "duplicate"
		end if 
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
	if (g_ExchangeVS == 'duplicate') {
		alert("이미 추가된 현장입니다.");
	}	
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