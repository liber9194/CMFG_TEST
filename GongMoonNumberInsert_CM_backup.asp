<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="misop.inc" -->
<%

Seq = Request("Seq")
Page = Request("Page")
type__ = Request("type__")

db_id = session("db_id")

c_date = year(Date)

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
		'Max_Num_End = 15000
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
	  ' if isnull(DbRec("Mnumber")) then
		  ' Max_Num_End = 15001
	  ' else
		  ' Max_Num_End = DbRec("Mnumber") + 1	
		  ' 'if Max_Num_End > 999 then
			' 'Max_Num_End = 1
		  ' 'end if
	  ' end if 
	  
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
	//parent.frames[2].location = "right_main_gongmoon_CM.asp?Page=" + CPage + "&type__=" + Ctype ; 접수함에서 승인처리 안함
	parent.frames[2].location = "right_main_gongmoon_insert_CM.asp?Page=" + CPage + "&type__=" + Ctype ; // 발송함에서 기안자가 직접 따는 것으로 변경(2017.11.24/경익수 상무 요청)
}
-->
</script>

</HEAD>
<!--BODY onload="PrintFrm()"-->
<BODY onload="PrintFrm()">



</BODY>
</HTML>