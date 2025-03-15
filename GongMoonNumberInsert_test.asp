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
  
sql="select max(number_) as Mnumber from result_tbl where year_ = '" & c_date & "' and type__='" & type__ & "' "
	
DbRec.Open sql, DbTestCon

if DbRec.EOF or DbRec.BOF then
	
	NoData = True

	if type__ = "" then	
		Max_Num_End = 7000
	elseif type__ = "1" then
		Max_Num_End = 1
	elseif type__ = "2" then
		Max_Num_End = 10001
	elseif type__ = "3" then
		Max_Num_End = 15000
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
		  Max_Num_End = 15001
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
set Result = DbTestCon.execute(sqlstr)

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
	parent.frames[2].location = "right_main_gongmoon_test.asp?Page=" + CPage + "&type__=" + Ctype ;
}
-->
</script>

</HEAD>
<!--BODY onload="PrintFrm()"-->
<BODY onload="PrintFrm()">



</BODY>
</HTML>