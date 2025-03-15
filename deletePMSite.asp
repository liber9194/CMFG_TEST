
<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<%

	i_userid = request("i_userid")
	site_code = request("site_code")
	

	IF i_userid <> "" AND  site_code <> "" Then 

	    '마지막 history 검색
		sql = "  Select TOP 1 CNT_ From PM_manager_His where UserId = '" & i_userid & "' and site_code='" & site_code & "' ORDER BY CNT_ DESC "
		Set rs=Server.CreateObject("ADODB.Recordset")
		rs.CursorType=1 ' CursorTyp : 0 : adOpenForwardOnly (기본값) 	1 : adOpenKeyset 	2 : adOpenDynamic 	3 : adOpenStatic 
		c_CNT = ""
		rs.Open sql, DbCon
		If rs.Recordcount <> 0 Then 
			cnt = rs.Recordcount
			c_CNT = rs("CNT_")
			'프로젝트 삭제, 프로젝트 history update
			sql = "DELETE PM_Manager WHERE UserId = '" & i_userid & "' and site_code='" & site_code & "';" & vbCrLf
			sql = sql & " IF EXISTS( SELECT  * " & _
                                    "           FROM    PM_Manager_His " & _
                                    "           WHERE   Site_Code = '" & site_code & "' AND UserID = '" & i_userid & "') " & _
                                    "   UPDATE  PM_Manager_His " & _
                                    "   SET     OUT_DATE = convert(varchar(10),getdate(),120) " & _
                                    "   WHERE   Site_Code = '" & site_code & "' AND UserID = '" & i_userid & "' AND CNT_ = " & c_CNT & " " & _
                                    "ELSE " & _
                                    "   INSERT INTO PM_Manager_His(Site_Code,UserID,IN_DATE,OUT_DATE,CNT_) VALUES ('" & site_code & "','" & i_userid & "','',convert(varchar(10),getdate(),120),(SELECT  isnull(MAX(CNT_),0)+1 AS Cnt__ FROM  PM_Manager_His   WHERE   Site_Code = '" & site_code & "')); " & vbCrLf
									
			'response.write sql
			'response.end
			Set Result = DbCon.execute(sql)
			Set Result=Nothing
			DbCon.close
			m_sus = "ok"
		else
			m_sus = "not"
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
	if (g_ExchangeVS == 'not') {
		alert("삭제할 현장이 없습니다. 관리자에게 문의하십시오.");
	}	
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