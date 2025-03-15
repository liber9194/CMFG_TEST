<%@ LANGUAGE="VBSCRIPT"%>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<!-- #include file="../Mis/WebWrite/config.asp" -->

<%
		
Site_Code = Request("Site_Code")
index = Request("index_")
db_id = session("db_id")

isErr = "N"

IF Site_Code = "" or index = "" or db_id = "" then
	response.write "<script language='javascript'>"
	response.write "alert('세션이 만료되었습니다. 다시 접속 후 저장해주세요.');"
	response.write "window.parent.parent.close();"
	response.write "</script>"
	response.end
end if 


sql = " INSERT TBL_Appointment_Header(Site_Code, WorkID, Create_User, Create_Date, Edit_User, Edit_Date, Main_Work_YN, Sub_Work_NM) "
sql = sql & " VALUES('" & Site_Code & "', (select isnull(max(WorkID)+1,2) from TBL_Appointment_Header where Site_Code ='" & Site_Code & "' AND Main_Work_YN = 'N'), '" & db_id & "', getdate(), NULL, NULL, 'N', (select prj_sub_saup_name from Prj_Main_Other_01 WHERE Site_Code = '" & Site_Code & "'	AND index_ = " & index & ")) "

response.write sql
DbCon.execute(sql)

if DbCon.Errors.count > 0 Then
	isErr = "Y"
end if	

%>


<HTML>
<HEAD>
<TITLE>Save</TITLE>
<LINK href="../../Home/css/default_ver_up.css" type="text/css" rel="stylesheet">
<link	REL="stylesheet" TYPE="text/css"	HREF="../../Home/css/bootstrap.min.css">

<script language="JavaScript" src="../../../../js_common/jquery-3.1.1.min.js"></script>
<script language="JavaScript" src="../../../../js_common/bootstrap.min.js"></script>
<script language="JavaScript">
$( document ).ready(function() {
	if ("<%=isErr%>" == "Y") {
		alert('공사 생성 중 오류 발생\nICT팀에 문의하세요.');
		return;
	} else {
	
		alert('공사 생성 완료되었습니다.')
		window.parent.opener.location.reload();
		window.parent.close();
	}	
});
</script>

</HEAD>
<BODY>
</BODY>
</HTML>