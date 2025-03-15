<%@ Language=VBScript %>
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="misop.inc" -->
<%
Seq = Request("Seq")
Page = Request("Page")
type__ = Request("type__")
Resident_chef = Request("Resident_chef") '상주 책임이 공문번호 생성한 경우 Y, 그 외 해당 변수 값 없음(23.01.19 경익수 상무 요청)

db_id = session("db_id")

c_date = year(Date)

'기 생성된 공문번호 존재하는지 체크(23.01.19)
sql = " select cnt = count(*) "
sql = sql & " from result_tbl "
sql = sql & " where o_seq = " & Seq

Set Db_Number_Chk = Server.CreateObject("ADODB.Recordset")
Db_Number_Chk.CursorType=1

Db_Number_Chk.Open sql, DbCon

if Db_Number_Chk.Recordcount <> 0 then	

' response.write "<script language='javascript'>"
' response.write "alert('" & Db_Number_Chk("cnt") & "');"
' response.write "history.go(-1);"
' response.write "</script>"
' response.end

	cnt = Db_Number_Chk("cnt")
	
	response.write "<script language='javascript'>"
	response.write "alert('" & cnt & "');"
	response.write "</script>"
	
	cnt = "1"
	
	
	'if Cint(Db_Number_Chk("cnt")) > 0 then	'기 생성된 공문번호 존재하면
	'if Db_Number_Chk("cnt") <> "0" then	'기 생성된 공문번호 존재하면
	if Cint(cnt) > 0 then	'기 생성된 공문번호 존재하면
		response.write "<script language='javascript'>"
		response.write "alert('공문번호가 이미 존재합니다.\n페이지를 새로고침하세요.');"
		response.write "history.go(-1);"
		response.write "</script>"
		response.end
	end if
end if

Set DbRec=Server.CreateObject("ADODB.Recordset")
DbRec.CursorType=1

'type__ = "" 과 type__ = "3"은 공문번호 생성을 같이 함 (7000번대~)
if type__ = "" or type__ = "3" then
	'cond = " and (type__ = '' or type__ = '3') and (type__ = '' or (type__ = '3' and result_date > '2018-02-12')) "
	cond = " and (type__ = '' or type__ = '3') and (type__ = '' or (type__ = '3' and result_date >= '2018-02-12' and number_ < 15010)) "
else
	cond = " and type__ = '" & type__ & "' "
end if
  
'sql="select max(number_) as Mnumber from result_tbl where year_ = '" & c_date & "' and type__='" & type__ & "' "
sql="select max(number_) as Mnumber from result_tbl where year_ = '" & c_date & "' " & cond
	
DbRec.Open sql, DbCon

if DbRec.EOF or DbRec.BOF then
	
	NoData = True

	if type__ = "" then	
		Max_Num_End = 7000
	elseif type__ = "1" then
		Max_Num_End = 1
	elseif type__ = "2" then
		Max_Num_End = 10001
	elseif type__ = "3" then
		Max_Num_End = 7000
	end if
else
	
	NoData = False


	if type__ = "" then	
	  if isnull(DbRec("Mnumber")) then
		  Max_Num_End = 7001
	  else
		  Max_Num_End = DbRec("Mnumber") + 1	
		  'if Max_Num_End > 999 then
			'Max_Num_End = 1
		  'end if
	  end if
	elseif type__ = "1" then
	  if isnull(DbRec("Mnumber")) then
		  Max_Num_End = 1
	  else
		  Max_Num_End = DbRec("Mnumber") + 1	
		  'if Max_Num_End > 999 then
			'Max_Num_End = 1
		  'end if
	  end if
	elseif type__ = "2" then
	  if isnull(DbRec("Mnumber")) then
		  Max_Num_End = 10001
	  else
		  Max_Num_End = DbRec("Mnumber") + 1	
		  'if Max_Num_End > 999 then
			'Max_Num_End = 1
		  'end if
	  end if
	elseif type__ = "3" then
	  if isnull(DbRec("Mnumber")) then
		  Max_Num_End = 7001
	  else
		  Max_Num_End = DbRec("Mnumber") + 1	
		  'if Max_Num_End > 999 then
			'Max_Num_End = 1
		  'end if
	  end if
	end if


end if






sqlstr = " insert result_tbl([o_seq], [number_], [year_], [sabun], [result_date],type__) values("
sqlstr = sqlstr & "" & Seq & ","
sqlstr = sqlstr & "" & Max_Num_End & ","
sqlstr = sqlstr & "'" & c_date & "',"
sqlstr = sqlstr & "'" & db_id & "',"
sqlstr = sqlstr & "'" & Date & "',"
sqlstr = sqlstr & "'" & type__ & "')"
set Result = DbCon.execute(sqlstr)

set Result = nothing









'response.write Passdate

'response.write Stastarting

'response.write Staterminal

'response.write Railkbn

%>

<HTML>
<HEAD>
<TITLE>Del</TITLE>


<script language="JavaScript">
<!--
function PrintFrm(){
//	window.close();
//	opener.window.location.href="Right_Main_GongMoon.asp";
//	parent.right.location.href="Rail_List.asp";
//}
	//var firstWin = window.parent.opener;
	//firstWin.location = "right_main_gongmoon.asp?Seq=" + pItemBoardID + "&Page=" + <%=Page%>;
	//alert("t");
	//window.close();


	//var CPage = "<%=Page%>"
//alert(CPage);
	//window.location.href = "right_main_gongmoon.asp?Seq=" + pItemBoardID + "&Page=" + CPage;
	
	
	var CPage = "<%=Page%>"	
	var Ctype = "<%=type__%>"
	parent.frames[2].location = "right_main_gongmoon_ver_up.asp?Page=" + CPage + "&type__=" + Ctype ;
}
-->
</script>

</HEAD>
<!--BODY onload="PrintFrm()"-->
<BODY onload="PrintFrm()">



</BODY>
</HTML>