<%@ Language=VBScript CODEPAGE="65001" %>
<!--#include file="adovbs_Basic.inc" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<!--#include file="../../../../default_properties.asp"-->
<%
	db_id = session("db_id")
	chk_obj = Request("ccBox[]") 'classic asp 는 체크된 체크박스 배열만 넘겨받음
	chk_list = Split(chk_obj, ",")
	
	dim fs
	set fs = Server.CreateObject("Scripting.FileSystemObject")
	
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.CursorType=1 ' CursorTyp : 0 : adOpenForwardOnly (기본값) 	1 : adOpenKeyset 	2 : adOpenDynamic 	3 : adOpenStatic 		
	
	For i=0 to UBound(chk_list) ' *** 참고 : 배열의 인덱스는 0부터 시작, UBound는 배열의 마지막 인덱스가 나오기 때문에, 배열 크기 구할 땐 UBound + 1
		
		'공지 관련 폴더, DB 삭제
		sql = " SELECT * "
		sql = sql & " FROM TBL_NOTICE "
		sql = sql & " WHERE NOTICE_ID = '" & Replace(chk_list(i), " ", "") & "' "		
		
		rs.Open sql, DbCon
		 
		if not rs.EOF then
			
			Bpath = rs("NOTICE_HTML_PATH")
			physicalPath = g_file_real_path & "\Notice\" & rs("NOTICE_HTML_PATH")
			fullPath = g_file_path & "/Notice/" & Bpath

			if (fs.FolderExists(physicalPath)) Then		'같은이름의 파일이 있을 때
				fs.DeleteFolder(physicalPath)				
			end if
				
			sql = " delete TBL_NOTICE_FILE "
			sql = sql & " where NOTICE_ID = '" & Replace(chk_list(i), " ", "") & "' "
			
			sql = sql & " delete TBL_NOTICE "
			sql = sql & " where NOTICE_ID = '" & Replace(chk_list(i), " ", "") & "' "				
			
			Set Result = DbCon.execute(sql)
			Set Result=Nothing
			
		end if
		
		rs.Close
		
	Next

	Set rs=Nothing
	Set fs=Nothing
%>
<HTML>
<HEAD>
<TITLE>Save</TITLE>
<script language="JavaScript">
function fnMessage(){
	alert("삭제되었습니다.");
	parent.goSearch();
}
</script>

</HEAD>
<BODY onload="fnMessage()">
</BODY>
</HTML>