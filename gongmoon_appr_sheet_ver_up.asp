<%@ LANGUAGE="VBSCRIPT" %>

<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=11">	

<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="misop.inc" -->
<!-- #include file="../../../../default_properties.asp" -->
<%
'인증체크 추가 2016.04.26
IF session("db_id") = "" Then
	'Response.Redirect "http://sfg.dohwa.co.kr/"
	Response.Redirect g_home_url
End IF


'페이지 접속로그 추가 2016.04.21==================================================

	strUserIP  = Request.ServerVariables("REMOTE_HOST")	'로그인 IP 기록
	strSql = " INSERT INTO PAGE_LOG_INFO([IP],[EMP_ID],[EMP_NAME],[PAGE_NAME],[PAGE_ACTION]) "
    strSql = strSql &   " VALUES('" & strUserIP & "'"
	strSql = strSql &   " ,'" & db_id & "'"
	strSql = strSql &   " ,'" & db_name & "' "
	strSql = strSql &   " ,'gongmoon_appr_sheet_ver_up.asp' "
	strSql = strSql &   " ,'공문결재현황 목록'"
	strSql = strSql &   " ) "

	Set Result = DbCon.execute(strSql)
	Set Result=Nothing
	
'페이지 접속로그 추가 2016.04.21==================================================

number		= Request("number")


QSelect  = Request("QSelect")

Qgubun   = Request("Qgubun")

type__  = Request("type__")

db_id 	 	= session("db_id")
db_level 	 	= session("db_level")

site_code 	 	= session("site_code")

search = Request("search")
search_txt = Request("search_txt")

%>

<title>mail_list</title>
<meta name="GENERATOR" Content="Microsoft Visual Studio 7.0">
<meta name="CODE_LANGUAGE" Content="C#">
<meta name="vs_defaultClientScript" content="JavaScript">

<link rel="stylesheet" href="../../Home/css/default_ver_up.css" type="text/css">
<link	REL="stylesheet" TYPE="text/css"	HREF="../../Home/css/bootstrap.min.css">
<style type="text/css">
.td_title{
 	border:1px solid #dddddd; 
	background-color:#EFEFEF;
	height:25px;
	text-align: center;
}

.td_header{
 	border:1px solid #dddddd; 
	background-color:#EFEFEF;
	height:25px;
	text-align: center;
	font-weight: bold;
}
.td_data{
	border:1px solid #dddddd; 
	height:25px;
	text-align: left;
	padding-left: 7px;
	background-color:#ffffff;
}
.divModify {
    border-top:1px solid #c9dae4;
	border-left:1px solid #c9dae4; 
 	border-right:1px solid #c9dae4; 
	border-bottom:1px solid #c9dae4;
	background-color:#f7fcff;
    padding:7px;
    overflow: auto;
    width: 520px;
    height:198px;	
}
.popuplist tr {
	height:30px;
}
.popuplist th, td {
	text-align:center;
}
</style>
<link	REL="stylesheet" TYPE="text/css"	HREF="./css/admin.css">
<script language="JScript" src="../ezEmail/lang/ezEmail_ko.js"></script>
<script language="JavaScript" src="../../Home/myoffice/common/mouseeffect.js"></script>
<STYLE> 
P { MARGIN-BOTTOM: 0mm; MARGIN-TOP: 0mm } 
</STYLE>

<!--<script language="JScript" src="../ezEmail/js/emails.js"></script>-->
<script language="JScript" src="../ezEmail/js/email_tree.js"></script>
<script language="JScript" src="../ezEmail/js/string_component.js"></script>

<script>
function goToPage1()
{
	var search_txt = document.all.search_txt.value;
	search_txt = search_txt.replace(/\s/gi, ""); 
	
	window.location.href = "gongmoon_appr_sheet_ver_up.asp?search=" + document.all.search.value + "&search_txt=" + search_txt;
}
</script>

</HEAD>
<!--<body style="BEHAVIOR:url('#default#userData');OVERFLOW:hidden" id="theBody" class="mainbody">-->
<body style="BEHAVIOR:url('#default#userData');OVERFLOW-y:auto;" id="theBody" class="mainbody">

	<table class="layout" style="width: 60%;">

		<tr>
			<% if db_level = "S" OR db_level = "Z" THEN %>		
				<td valign="top" height="40"><h1>PM 결재현황</h1>
			<% elseif db_level = "P" then %>
				<td valign="top" height="40"><h1>공문결재현황</h1>
			<% END IF %>		

		
		<!-- 본문 시작 -->
		<% if db_level = "S" OR db_level = "Z" THEN %>	
		
			<div class="row" style="padding-top: 43px;padding-bottom:5px">
				<div class="col-sm-6" style="text-align:left;">
					<div class="form-group form-group-jh">
						<div class="col-sm-3" style="padding-left:0px;padding-right:0px;">
							<select class="form-control input-sm" name="search">       
								<option VALUE="pm_id" <% if search = "pm_id" then response.write "selected" end if %>>PM사번</option>
								<option VALUE="pm_name" <% if search = "pm_name" then response.write "selected" end if %>>PM성명</option>				
							</select>
						</div>
						<div class="col-sm-4" style="padding-left:5px;padding-right:0px;">		
							<input class='form-control' type='textbox' size="15" name='search_txt' VALUE="<%=search_txt%>" onkeypress="javascript : if (event.keyCode == 13) goToPage1();">
						</div>
						<div class="col-sm-5" style="padding-left:5px;padding-right:0px;">							
							<button type="button" class="btn btn-default btn-jh" onClick="goToPage1();">검색</button>
						</div>						
					</div>
				</div>
				<div class="col-sm-6" style="text-align:left;">
				</div>
			</div>
			
			<div style="text-align:left;">
				<h2 style="font-size:10pt;margin-top:0px;">* 결재된 문서 : 후결재/결재올림 처리한 문서</h2>
			</div>
		
			<%
			Set rs=Server.CreateObject("ADODB.Recordset")
			rs.CursorType=1
		
			sql = " select a.c_userid, c_name = replace(a.c_name, ' ', ''), pm = replace(a.c_name, ' ', '') + ' (' + a.c_userid + ')' "
			
			if Request.ServerVariables("http_host") = g_domain then
				sql = sql & " from [" & g_cmfgDB & "].cug_Test.dbo.user_tbl a "
			else
				sql = sql & " from sfg.cug_Test.dbo.user_tbl a "
			end if
						
			sql = sql & " 	inner join dh_sap.dbo.tbl_sap_emp_info b on a.c_userid = b.emp_id "
			sql = sql & " where a.c_level = 'P' "
			sql = sql & " 	and b.RETIRE_DT > (select format(getdate(), 'yyyyMMdd')) "
			
			if search_txt <> "" then
				if search = "pm_id" then
					sql = sql & " and a.c_userid = '" & search_txt & "' "
				elseif search = "pm_name" then
					sql = sql & " and replace(a.c_name, ' ', '') = '" & search_txt & "' "
				end if
			end if
			
			rs.Open sql, DbCon_Mis
			%>			
			<table class="mainlist" id ='test'>
				<tr>
					<th style="width:5%;text-align:center;">No</th>
					<th style="width:20%;text-align:center;">해당PM</th>
					<th style="width:10%;text-align:right;">접수문서(건)</th>
					<th style="width:10%;text-align:right;">읽음 문서(건)</th>
					<th style="width:10%;text-align:right;">안읽음 문서(건)</th>
					<th style="width:10%;text-align:right;">결재된 문서(건)</th>
				</tr>
				
			<%		
			if rs.Recordcount <> 0 then
				cnt = 0
			%>				
				<%
				for i = 0 to rs.Recordcount - 1
				
					Set DbRec=Server.CreateObject("ADODB.Recordset")
					DbRec.CursorType=1
		
					sql = "			select count(*) as 'receive_cnt', "
					sql = sql & " 		count(case when b.o_visited > 0 then 1 end) as 'read_cnt', "
					sql = sql & " 		count(case when b.o_visited = 0 then 1 end) as 'unread_cnt', "
					sql = sql & " 		count(case when b.Result_Add = '1' or b.Result_Add = '3' then 1 end) as 'appr_cnt' " '후결재, 결재올림을 결재된 문서로 인정
					sql = sql & " 	from office_tbl a "
					sql = sql & " 		inner join receive_tbl b on a.o_seq = b.o_seq and b.receive_del = '' "
					sql = sql & " 	where b.o_receive_id = '" & rs("c_userid") & "' "
					sql = sql & " 		and a.office_del = '' "
					sql = sql & " 		and b.o_rdel_flag = 0 "
			
					DbRec.Open sql, DbCon
					
					if DbRec.Recordcount <> 0 then
						cnt = cnt + 1
				%>			
						<tr height="25">
							<td style="text-align:center;padding:5px;"><%=cnt%></td>
							<td style="text-align:center;padding:5px;"><%=rs("pm")%></td>
							<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("receive_cnt"),0,-1,0,-1)%></td>
							<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("read_cnt"),0,-1,0,-1)%></td>
							<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("unread_cnt"),0,-1,0,-1)%></td>
							<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("appr_cnt"),0,-1,0,-1)%></td>
						</tr>
					<% end if %>
					<% set DbRec = Nothing %>
				<% rs.MoveNext %>
				<% next %>			
			<% end if %>
			<% set rs = Nothing %>			
		
			</table>			
		<% elseif db_level = "P" then %>
			<%
			
			Set DbRec=Server.CreateObject("ADODB.Recordset")
			DbRec.CursorType=1
		
			sql = "			select count(*) as 'receive_cnt', "
			sql = sql & " 		count(case when b.o_visited > 0 then 1 end) as 'read_cnt', "
			sql = sql & " 		count(case when b.o_visited = 0 then 1 end) as 'unread_cnt', "
			sql = sql & " 		count(case when b.Result_Add = '1' or b.Result_Add = '3' then 1 end) as 'appr_cnt' "
			sql = sql & " 	from office_tbl a "
			sql = sql & " 		inner join receive_tbl b on a.o_seq = b.o_seq and b.receive_del = '' "
			sql = sql & " 	where b.o_receive_id = '" & db_id & "' "
			sql = sql & " 		and a.office_del = '' "
			sql = sql & " 		and b.o_rdel_flag = 0 "
			
			DbRec.Open sql, DbCon

			if DbRec.Recordcount <> 0 then
			%>
				<table class="mainlist" id ='test' style="margin-top:45px;">
					<tr>
						<th style="width:25%;text-align:center;">접수문서(건)</th>
						<th style="width:25%;text-align:center;">읽음 문서(건)</th>
						<th style="width:25%;text-align:center;">안읽음 문서(건)</th>
						<th style="width:25%;text-align:center;">결재된 문서(건)</th>
					</tr>
					<tr height="25">
						<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("receive_cnt"),0,-1,0,-1)%></td>
						<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("read_cnt"),0,-1,0,-1)%></td>
						<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("unread_cnt"),0,-1,0,-1)%></td>
						<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("appr_cnt"),0,-1,0,-1)%></td>
					</tr>			
				</table>
			<% end if %>
		<% END IF %>
		<!-- 본문 끝 -->
		
		</td>
	</tr>
	</table>
	
	<% set DbRec = Nothing %>
</body>
</HTML>



