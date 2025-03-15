<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="../dbopen.asp"-->
<!--#include file="../ReqVariant.asp"-->
<%

number		= Request("number")


if Request("page")="" then
	curpage=1
else
	curpage=cint(Request("page"))
end if

if Request("startpage")="" then
	startpage=1
else
	startpage=cint(Request("startpage"))
end if

ipp=15
ten=5




Set DbRec=Server.CreateObject("ADODB.Recordset")
DbRec.CursorType=1


str = "SELECT * FROM addr_tbl WHERE reg_id = '" & db_id & "' ORDER BY a_seq ASC"
DbRec.Open str, DbCon


if DbRec.Recordcount <> 0 then

	DbRec.MoveLast
	postcount=DbRec.Recordcount

	totpage=int(postcount/ipp)
	totpage=cint(totpage)

else

	postcount = 0
	totpage = 0

end if

if(totpage * ipp) <> postcount then totpage = totpage + 1

For a=1 to (curpage-1) * ipp
	DbRec.MovePrevious
Next 

pg=Request.QueryString("page")
if isEmpty(pg) then
	pg=1
else
	pg=pg+0
end if

if pg<1 then
	pg=1
end if


if DbRec.Recordcount <> 0 then


	sql="SELECT Count(*) as totalcount FROM addr_TBL WHERE reg_id = '" & db_id & "'"

	Set rs=DbCon.Execute(sql)

	lastpg=1+Int((rs("totalcount")-1)/ipp)
	if pg>lastpg then
	pg=lastpg
	end if

	nextpg=pg+1
	prevpg=pg-1
	endpg=pg*ipp
	startpg=endpg-ipp+1
else
	lastpg = 1
	pg = 1
	nextpg = 2
	prevpg = 0
	endpg = 15
	startpg = 1
	
end if

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>사내 게시판(도화종합기술공사)</title>

<style>
   #dami:hover {color:Black;font-weight: bolder}
    SPAN { cursor:hand; }
   #RED:hover {color:RED}
</style>

<style type="text/css">
<!--
        A:link {font: 12 굴림체,Arial;  COLOR: "#0000A0"; text-decoration: none;}
				A:active {font: 12 굴림체,Arial;  COLOR: "#0000A0"; text-decoration: none;}
        A:visited {font: 12 굴림체,Arial; COLOR: "#0000A0"; text-decoration: none;}
-->
</style>


<SCRIPT Language="JavaScript">
</SCRIPT>

</head>


<body bgcolor="#FFF0C6">
<!-- Insert HTML here -->

<!-- 등록된 사용자 ID와 비밀번호가 정확한 경우 -->

<br>
<div align="center"><center>
<table border=0 align=center width=400 bgcolor=#DEDEC0 cellpadding=5>
	<tr><td align=center><font face=굴림><b>주소록</b></font></td>
	</tr>
</table></div>
<br><br>


<div align="center"><center>
<table border="0" cellpadding="3" cellspacing="1" width="500" bgcolor="#FFFFFF">
  <tr><!--#EBEBD8, 003399-->
	<td align="center"  bgcolor="#000000" width="50">
			<font face="굴림" color="#ECECFF"><b><DIV STYLE="font:12;">번</font>
			<font face="굴림" color="#DFDFFF">호</DIV></b></font></td>

    	<td align="center"  bgcolor="#000000" width="100">
			<font face="굴림" color="#ECECFF"><b><DIV STYLE="font:12;">주</font>
			<font face="굴림" color="#DFDFFF">소</font>
			<font face="굴림" color="#C1FFFF">록</font>
			<font face="굴림" color="#BBFFFF">이</font>
			<font face="굴림" color="#ECECFF">름</DIV></b></font></td>

    	<td align="center"  bgcolor="#000000" width="200">
			<font face="굴림" color="#ECECFF"><b><DIV STYLE="font:12;">설</font>
			<font face="굴림" color="#DFDFFF">명</DIV></b></font></td>

    	<td align="center"  bgcolor="#000000" width="50">
			<font face="굴림" color="#ECECFF"><b><DIV STYLE="font:12;">개</font>
			<font face="굴림" color="#C1FFFF">수</DIV></b></font></td>

  </tr>

<% if postcount <> 0 then %>
  <font color="#000000" >

<%
	For i = 1 to ipp
		if totpage = curpage then
			value = postcount Mod ipp
			if i > value and value <> 0 then
				Exit For
			end if
		end if
%>

</font>
  <tr onMouseOver = "this.style.backgroundColor = '#F9F9F9'" onMouseOut="this.style.backgroundColor='#F4F4EA'"><!--#F4F4EA -->
    	<td bgcolor="#FFFADC" align="right"><DIV STYLE="font:12;">

		<font face="굴림"  color="#808080"><%=DbRec("a_seq")%></font></DIV></td>

    	<td bgcolor="#F4F4EA"><DIV STYLE="font:12;">

		<a  href="content.asp?number=<%=DbRec("a_seq")%>&page=<%=curpage%>&startpage=<%=startpage%>&db_id=<%=db_id%>&db_passwd=<%=db_passwd%>&db_name=<%=db_name%>&db_email=<%=db_email%>&db_level=<%=db_level%>&db_longname=<%=db_longname%>&db_acc=<%=db_acc%>"  id="dami">
			<font color="#000080" face="굴림"><%=DbRec("a_name")%></font></a></DIV></td>

    	<td bgcolor="#FFFADC"><DIV STYLE="font:12;">
		<font face="굴림" color="#808080"><%=DbRec("a_desc")%></font></DIV></td>

    	<td bgcolor="#F4F4EA"><DIV STYLE="font:12;">
		<font face="굴림" color="#808080"><%=DbRec("a_count")%></font></DIV></td>

  </tr>
  <font color="#000000" ><%
DbRec.MovePrevious
Next
%>
</font>

<% end if %>

</table>
</center></div>


<div align="center">
<table border="0"  cellpadding="3" cellspacing="0" width="500">

<tr><td align="left" width="150">



	<font color="#000000" > <%if prevpg<1 then%> </font>
	<img src="../images/ge-prev.gif" alt="이전 페이지로" border=0>
	<font color="#000000" >  <%else%> </font>
	<a href="board.asp?page=<%=prevpg%>&db_id=<%=db_id%>&db_passwd=<%=db_passwd%>&db_name=<%=db_name%>&db_email=<%=db_email%>&db_level=<%=db_level%>&db_longname=<%=db_longname%>&db_acc=<%=db_acc%>" id="RED">
	<img src="../images/ge-prev.gif" alt="이전 페이지로" border=0></a>
	<font color="#000000" >  <%end if%> <%if nextpg > lastpg then%> </font>
	<img src="../images/ge-next.gif" alt="다음 페이지로" border=0>
	<font color="#000000" >  <%else%> </font>
	<a href="board.asp?page=<%=nextpg%>&db_id=<%=db_id%>&db_passwd=<%=db_passwd%>&db_name=<%=db_name%>&db_email=<%=db_email%>&db_level=<%=db_level%>&db_longname=<%=db_longname%>&db_acc=<%=db_acc%>" id="RED">
	<img src="../images/ge-next.gif" alt="다음 페이지로" border=0></a>
	<font color="#000000" >  <%end if%> 
	<a href="write.asp?db_id=<%=db_id%>&db_passwd=<%=db_passwd%>&db_name=<%=db_name%>&db_email=<%=db_email%>&db_level=<%=db_level%>&db_longname=<%=db_longname%>&db_acc=<%=db_acc%>" 
			onmouseover="window.status=('글을 등록하기');return true;"  
			onmouseout="window.status=('&nbsp;');return true;" id="RED"> 
	<img src="../images/ge-write.gif" alt="나도 한마디" border=0></a>

</td>
<td align="right"><DIV STYLE="font:12;">

<font color="#000000" >
	<%if totpage>ten then%> 
		<%if startpage=1 then%> [ </font>
		<font color="#808000"  face="굴림">이전 5페이지</font>
		<font color="#000000" > ] 
		<%else%> [<a href="board.asp?page=<%=cint(startpage)-ten%>&startpage=<%=cint(startpage)-ten%>&db_id=<%=db_id%>&db_passwd=<%=db_passwd%>&db_name=<%=db_name%>&db_email=<%=db_email%>&db_level=<%=db_level%>&db_longname=<%=db_longname%>&db_acc=<%=db_acc%>" id="RED"> 이전 5 페이지 </a>] 
		<%end if%> 
		<%
		For a=startpage to startpage+ten-1
		if a>totpage then
		exit for
		else
		if a=curpage then
		%>
		<font color="#ff0000" > <%=a%> </font>
		<%else%> <a href="board.asp?page=<%=a%>&startpage=<%=startpage%>&db_id=<%=db_id%>&db_passwd=<%=db_passwd%>&db_name=<%=db_name%>&db_email=<%=db_email%>&db_level=<%=db_level%>&db_longname=<%=db_longname%>&db_acc=<%=db_acc%>" id="RED"><%=a%></a> 
		<%End if%> 
		<%end if%> 
		<%Next%> 
		<%if((startpage\ten)=(totpage\ten)) then%> [ </font>
		<font color="#808000"  face="굴림">다음 5페이지</font>
		<font color="#000000" > ] 
		<%else%>[<a href="board.asp?page=<%=a%>&startpage=<%=a%>&db_id=<%=db_id%>&db_passwd=<%=db_passwd%>&db_name=<%=db_name%>&db_email=<%=db_email%>&db_level=<%=db_level%>&db_longname=<%=db_longname%>&db_acc=<%=db_acc%>" id="RED"> 다음 5페이지 </a>] 
		<%end if%>
		<%else%>[ </font>
		<font color="#808000"  face="굴림">이전 5페이지</font>
		<font color="#000000" > ] 
		<%
		For a=startpage to totpage
		if a=curpage then
		%> <%=a%> 
		<%else%> 
		<a href="board.asp?page=<%=a%>&db_id=<%=db_id%>&db_passwd=<%=db_passwd%>&db_name=<%=db_name%>&db_email=<%=db_email%>&db_level=<%=db_level%>&db_longname=<%=db_longname%>&db_acc=<%=db_acc%>">
		<font color="#ff0000" ><%=a%></font></a> 
		<%end if
		next%> [ </font>
		<font color="#808000"  face="굴림">다음 5페이지</font>
		<font color="#000000" > ] 
	<%end if%></font></DIV>

</td></tr></table>


<hr size="1" width="500" color="#000080">
<div align="center"><center>

<table border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="1"></td>
  </tr>
  <tr>
    <td><DIV STYLE="font:12;"><font color="#000000" >Copyright by Dohwa Consulting Engineers co.ltd.<br> For more Information, 
    Contact <a href="mailto:cug@dohwa.co.kr">webmaster</a></font></DIV></td>
  </tr>
</table>
</center></div>

</DIV>
</body>
</html>
