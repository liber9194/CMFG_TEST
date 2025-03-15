<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="misop.inc" -->
<%
Seq = Request("Seq")


Set DbRec1=Server.CreateObject("ADODB.Recordset")
DbRec1.CursorType=1

	' -------- 새로운 공문 메시지 --------------
	'str = "select Count(*) as r_count  from office_tbl, receive_tbl "
	'str = str & " where o_visited = 0 and receive_tbl.o_receive_id = '207047' "
	'str = str & " and receive_tbl.o_seq = office_tbl.o_seq"

	'DbRec.Open str, DbCon
	'r_count = DbRec("r_count")

	'DbRec.Close
	' -------- 끝. 새로운 공문 메시지 --------------

if Seq <> "" then



	sqlstr 	= "select o_seq, o_savefile from save_file where o_seq = " & Seq & " "
	'Set Result1 = DbCon.execute(sqlstr)

	DbRec1.Open sqlstr, DbCon
'	html_ = DbRec1("o_savehtml")

	'if Result1.Recordcount > 0 then
	'	html_ = Result1("o_savehtml")
	'end if

	'Set Result=Nothing


	'sqlstr 	= "select o_seq, o_savefile from save_file where o_seq = " & Seq & " "
	'Set Result = DbCon.execute(sqlstr)
	
	'if Result.Recordcount > 0 then
		'file = Result("o_savehtml")
	'end if

	'Set Result=Nothing
			


end if

%>






<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<HTML>
<HEAD>
<title>mail_list</title>
<meta name="GENERATOR" Content="Microsoft Visual Studio 7.0">
<meta name="CODE_LANGUAGE" Content="C#">
<meta name="vs_defaultClientScript" content="JavaScript">
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<LINK rel="stylesheet" href="../../Home/css/default.css">

<STYLE> 
P { MARGIN-BOTTOM: 0mm; MARGIN-TOP: 0mm } 
</STYLE>


<!--body style="BEHAVIOR:url('#default#userData');OVERFLOW:hidden" id="theBody" class="mainbody"-->
<body class="mainbody1">

<% if DbRec1.Recordcount <> 0 then 

	For i = 1 to ipp

%>

<a href="downld.asp?file=<%Result1("o_savehtml")%>" id="white"
onmouseover="window.status=('첨부파일 다운로드하기(오른쪽 마우스 클릭하여 『다른이름으로 저장』을 선택');return true;"
onmouseout="window.status=('&nbsp;');return true;" ><%Result1("o_savehtml")%></a> </font>


	<%
	DbRec1.MovePrevious
	Next
	%>


<% end if %>

</body>
</HTML>



