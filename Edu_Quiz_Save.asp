
<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<%
	'db_id 	 	= session("db_id")
	db_id 	 	= Request("db_id")
	
	Edu_Year = Request("Edu_Year")
	Main_Chap = Request("Main_Chap")
	Sub_Chap = Request("Sub_Chap")

	if db_id = "" then
%>
<script language="javascript">
	alert("세션이 만료되었습니다.\n\n재로그인 후 이용하여 주십시오.");
	self.close();
</script>
<%	
		response.end
	end if

	'답 insert
	sql = ""
	
	For i = 1 To Request("answer").count		
		sql = sql & " INSERT Edu_Emp_Answer_TBL (Edu_Year, Emp_Num, Main_Chap, Sub_Chap, Q_No, Answer, Change_Date) "
		sql = sql & " VALUES(" & Edu_Year & ", '" & db_id & "', " & Main_Chap & ", " & Sub_Chap & ", " & i & ", '" & Replace(Request("answer")(i),"'","''") & "', getdate()) "
	next	
	
	if sql <> "" then
		Set Result = DbCon.execute(sql)
		Set Result=Nothing
	end if
	
	'직원이 퀴즈 최종완료했는지 확인 후 flag update
	sql = " SELECT Total_Q_Cnt = (SELECT SUM(Input_Count) FROM Edu_List_TBL WHERE Edu_Year = " & Edu_Year & ") "
	sql = sql & " ,Total_A_Cnt = (SELECT COUNT(*) FROM Edu_Emp_Answer_TBL WHERE Edu_YEAR = " & Edu_Year & " and Emp_Num = '" & db_id & "') "
	sql = sql & " ,Finish = case when (SELECT SUM(Input_Count) FROM Edu_List_TBL WHERE Edu_Year = " & Edu_Year & ") = (SELECT COUNT(*) FROM Edu_Emp_Answer_TBL WHERE Edu_YEAR = " & Edu_Year & " and Emp_Num = '" & db_id & "') then 'Y' else 'N' end "

	Set rs_Quiz_Chk = Server.CreateObject("ADODB.Recordset")
	rs_Quiz_Chk.CursorType=1
	
	rs_Quiz_Chk.Open sql, DbCon	
	
	if rs_Quiz_Chk("Finish") = "Y" then 
		sql = " 	  UPDATE Edu_Emp_List_TBL "
		sql = sql & " SET Finish_YN = 'Y', Change_Date = getdate() "
		sql = sql & " WHERE Edu_Year = " & Edu_Year & " and Emp_Num = '" & db_id & "' "
		
		Set Result = DbCon.execute(sql)
		Set Result=Nothing
	ELSE
		sql = " 	IF EXISTS(SELECT * FROM Edu_Emp_List_TBL WHERE Edu_Year = " & Edu_Year & " and Emp_Num = '" & db_id & "') "
		sql = sql & " 	BEGIN "
		sql = sql & " 		UPDATE Edu_Emp_List_TBL "
		sql = sql & " 		SET Change_Date = getdate() "
		sql = sql & " 		WHERE Edu_Year = " & Edu_Year & " and Emp_Num = '" & db_id & "' "
		sql = sql & " 	END "
		sql = sql & " ELSE "
		sql = sql & " 	BEGIN "
		sql = sql & " 		INSERT Edu_Emp_List_TBL (Edu_Year, Emp_Num, Finish_YN, Change_Date) "
		sql = sql & " 		VALUES(" & Edu_Year & ", '" & db_id & "', 'N', getdate()) "
		sql = sql & " 	END "
		
		Set Result = DbCon.execute(sql)
		Set Result=Nothing
	end if
%>
<HTML>
<HEAD>
<TITLE>Save</TITLE>
<script language="JavaScript">
function fnMessage(){
	alert("제출되었습니다.");
	//parent.goSearch();
		
	//window.location.href = "Popup_Edu_Quiz.asp?Edu_Year=<%=Edu_Year%>&Main_Chap=<%=Main_Chap%>&Sub_Chap=<%=Sub_Chap%>";
	
	window.opener.document.location.reload();	
	window.close();
	window.opener.close();
}
</script>

</HEAD>
<BODY onload="fnMessage()">
</BODY>
</HTML>