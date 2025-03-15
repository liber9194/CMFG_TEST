<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="misop.inc" -->
<%
'페이지 접속로그 추가 2016.04.21==================================================

	strUserIP  = Request.ServerVariables("REMOTE_HOST")	'로그인 IP 기록
	strSql = " INSERT INTO PAGE_LOG_INFO([IP],[EMP_ID],[EMP_NAME],[PAGE_NAME],[PAGE_ACTION]) "
    strSql = strSql &   " VALUES('" & strUserIP & "'"
	strSql = strSql &   " ,'" & db_id & "'"
	strSql = strSql &   " ,'" & db_name & "' "
	strSql = strSql &   " ,'ilban_gongmoon_read.asp' "
	strSql = strSql &   " ,'요청자료접수함 목록' "
	strSql = strSql &   " ) "

	Set Result = DbCon.execute(strSql)
	Set Result=Nothing
	
'페이지 접속로그 추가 2016.04.21==================================================

number		= Request("number")
QSelect		= Request("QSelect")
Qgubun		= Request("Qgubun")
type__		= Request("type__")

db_id 	 		= session("db_id")
db_level 	 	= session("db_level")
site_code 	 	= session("site_code")

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

if trim(Qgubun) <> "" then
	curpage=1
end if

ipp=15
ten=5


Set DbRec=Server.CreateObject("ADODB.Recordset")
DbRec.CursorType=1

	str = "SELECT A.o_seq, A.o_send_id, A.o_send_name, B.o_receive_id, A.o_subject, A.o_content, A.o_filename,a.private_name,private_number, "
	str = str & " A.o_filesize, A.o_doc_no, A.o_send_date, B.o_receive_date, A.o_send_longname, B.o_visited, "
	str = str & " ISNULL(C.number_,'') AS number_, ISNULL(C.year_,'') AS year_, ISNULL(C.sabun,'') AS sabun, ISNULL(C.result_date,'') AS result_date ,A.type__,[private__] "
	'str = str & " ,ISNULL(H.o_savehtml,'') AS savehtml_ "
	str = str & " ,ISNULL((select top 1 o_savehtml from user_html where o_seq = a.o_seq and o_receive_id = b.o_receive_id order by o_number desc), '') AS savehtml_ "
	str = str & " from office_tbl_i A "
	str = str & " inner join receive_tbl_i B ON A.o_seq = B.o_seq and B.receive_del = ''"
	str = str & " LEFT join result_tbl_i C ON A.o_seq = C.o_seq "

	'str = str & " LEFT join user_html H ON A.o_seq = H.o_seq and B.o_receive_id = H.o_receive_id " '현장에서 요청자료 여러번 보낼 경우 접수함에 리스트 중복으로 뜨는 것 방지

	'개인별로 하는

	'if db_level = "P" OR db_level = "Z" THEN
	if db_level = "Z" THEN
		str = str & " where B.o_receive_id = '" & db_id & "' "
	ELSE
		'str = str & " where B.o_receive_id = '" & site_code & "' "
		if db_level = "P" THEN
			site_code = "25-0000-000"
		END IF
		str = str & " where (B.o_receive_id = '" & site_code & "' or B.o_receive_id = '" & db_id & "') "

	end if

	str = str & " and A.type__ = '" & type__ & "' "
	str = str & " and A.office_del = '' "
'office_del = 'T'
	'if type__ = "" then
	'	str = str & " and A.type__ = '' "
	'else

	'end if

	str = str & " and B.o_rdel_flag = 0 "

	if QSelect = "제목" then
		str = str & " and A.o_subject Like '%" & Qgubun & "%' "
	end if
	if QSelect = "읽지않은공문" then
		str = str & " and B.o_visited < 1 "
	end if

	if QSelect = "보낸이" THEN
		str = str & " and A.o_send_name Like '" & Qgubun & "%' "
	END IF
	
	IF QSelect = "현장명" then
		str = str & " and A.o_send_longname Like '" & Qgubun & "%' "
	END IF

	str = str & " order by A.o_seq asc"	
	
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




cnt_i = 0

'DbRec.MoveLast

For a=1 to (curpage-1) * ipp

	'if (curpage-1) * ipp > DbRec.Recordcount then
		
	'else
	'	cnt_i = cnt_i + 1
		DbRec.MovePrevious
	'end if
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

	sql="SELECT Count(*) as totalcount from office_tbl_i, receive_tbl_i "
	sql = sql & " where (receive_tbl_i.o_receive_id = '" & db_id & "' ) and (office_tbl_i.o_seq = receive_tbl_i.o_seq)"
	sql = sql & " and (receive_tbl_i.o_rdel_flag = 0) "

	Set rs=DbCon.Execute(sql)

	lastpg=1+Int((rs("totalcount")-1)/ipp)
	if pg>lastpg then
	pg=lastpg
	end if

	nextpg=pg+1
	prevpg=pg-1
	endpg=pg*ipp
	startpg=endpg-ipp+1

	Nmod = DbRec.Recordcount mod 15
    Nanum = int(DbRec.Recordcount / 15)	

	if cnt_i > 0 then
		cnt_q = DbRec.Recordcount - Nanum * 15
	else
		cnt_q = 15
	end if 

	if Nmod <> 0 then
		Nanum = Nanum + 1
	end if

else

	lastpg = 1
	pg = 1
	nextpg = 2
	prevpg = 0
	endpg = 15
	startpg = 1
	
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
<script language="JScript" src="../ezEmail/lang/ezEmail_ko.js"></script>
<script language="JavaScript" src="../../Home/myoffice/common/mouseeffect.js"></script>
<STYLE> 
P { MARGIN-BOTTOM: 0mm; MARGIN-TOP: 0mm } 
</STYLE>



<script language="JScript" src="../ezEmail/js/emails.js"></script>
<script language="JScript" src="../ezEmail/js/email_tree.js"></script>
<script language="JScript" src="../ezEmail/js/string_component.js"></script>
<script>

var end_page = "<%=Nanum%>"


function initCheckBehavior1()
{
  var i, a;

  for (i = 0; i < document.links.length; ++i) {
    a = document.links[i];
    if (a.id.indexOf('UncheckAll_') != -1) {
      a.onclick = doCheckBehavior;
      a._CBNAME_ = a.id.substr(11) + '[]';
      a._CBCHECKED_ = false;
    }
    else if (a.id.indexOf('CheckAll_') != -1) {
      a.onclick = doCheckBehavior;
      a._CBNAME_ = a.id.substr(9) + '[]';
      a._CBCHECKED_ = true;
    }
  }
}

function doCheckBehavior()
{
  var i, cb = document.getElementsByName(this._CBNAME_);
  for (i = 0; i < cb.length; ++i) {
    cb[i].checked = this._CBCHECKED_;
  }
  return false;
}










function initCheckBehavior()
{

						var oColl = document.frmOutbox.ccBox;

						for (var i=0; i < document.frmOutbox.ccBox.length; i++) {
							oColl[i].checked = document.frmOutbox.cbox.checked;
					
						}


}










function Del_St(Cpage)
{

						if (confirm("공문 을 삭제 하시겠습니까?")) {      
								var oColl = document.frmOutbox.ccBox;
								var str = ''
								
								
								if ("<%=DbRec.Recordcount%>" == 1){
											
										if(document.frmOutbox.type_var.value != "") {
											if (oColl.checked==true){												
												if (document.frmOutbox.O_Del.value > 0) {
													alert("공문번호 가 발급된 공문은 삭제 할수 없습니다");
													str = '';
												} else {
													str += oColl.value + ';';
												}
											}		
											
											if (str == ''){
											}else {

												parent.frames[2].location  = "GongMoonDel_Office_i.asp?Page=" + Cpage + "&Str=" + str + "&type__=<%=type__%>";
											}
										}

								} else {
										
										for (var i=0; i < oColl.length; i++) {
																						
											if (oColl[i].checked==true){
												//alert(document.frmOutbox.type_var[i].value);
												if(document.frmOutbox.type_var[i].value != "") {
													//alert(document.frmOutbox.O_Del[i].value);
														if (document.frmOutbox.O_Del[i].value > 0) {
															alert("공문번호 가 발급된 공문은 삭제 할수 없습니다");
															str = '';
															break;
														}
														str += oColl[i].value + ';';
												
												}
											}

										}
										

										//alert(str);
										if (str == ''){
										}else {

											parent.frames[2].location  = "GongMoonDel_Office_i.asp?Page=" + Cpage + "&Str=" + str + "&type__=<%=type__%>";
										}
								}
						} 


}



		function goToPage(aaa)
		{

			var aaa1 = 0;

			if (aaa == "front") {

				aaa1 = parseInt(document.all.Cnum.value,0) - 1 
			} else if (aaa == "next") {
				aaa1 = parseInt(document.all.Cnum.value,0) + 1
			} else {
				aaa1 = document.all.Cnum.value
			}


				if (end_page < aaa1) {
					aaa1 = end_page

				}

				if (1 > aaa1) {
					aaa1 = 1
				}

				window.location.href = "ilban_gongmoon_read.asp?Page=" + aaa1 + "&type__=<%=type__%>" ;
 
		}





		function goToPage1(aaa1)
		{


				window.location.href = "ilban_gongmoon_read.asp?Page=" + aaa1 + "&QSelect=" + document.all.select.value + "&Qgubun=" + document.all.Gubun_i.value + "&type__=<%=type__%>"; 
		}






	function document.onselectstart()
	{
		event.cancelBubble = true;
		event.returnValue = false;
	}

		function SortPage(SortBy)
		{
			window.location.href = "../ezEmail/mail_read_i.asp?Seq=" + SortBy ;
		}



		function ItemRead_onclick(pItemBoardID,rID,Stype,HJname,visited,V1,qq_r,PP)
		{

			var e = event.srcElement;
			var eText = e.outerHTML;
			if(eText.substring(0,3)=="<B>"){
				e.outerHTML = eText.substring(3, eText.length);
			}
			
			var pheight = window.screen.availHeight;
			var pwidth = window.screen.availWidth;
			var pTop = (pheight - 720) / 2;
			var pLeft = (pwidth - 765) / 2;
			

			AutoCalc(qq_r);


			document.getElementById(V1).innerHTML="읽음";

			    window.open("../ezEmail/mail_read_i.asp?Seq=" + pItemBoardID + "&rID=" + rID + "&stype=" + Stype + "&HJname=" + HJname + "&visited=" + visited + "&PP=" + PP , "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,height=700,width=775,top=" + pTop + ",left=" + pLeft, "");	

		}



/**
		function checkBox_checked(pItemBoardID,obj,Cpage)
		{

			if (obj.checked == true) {

						if (confirm("공문번호 를 생성 하시겠습니까?")) {      

								parent.frames[2].location  = "GongMoonNumberInsert.asp?Seq=" + pItemBoardID + "&Page=" + Cpage + "&type__=<%=type__%>";

								f_submit();

								return true;
						} else {

							obj.checked = 0;		
						}
			} else {
				
				obj.checked = 1;

			}

		}
**/



function new_mail_onclick1() 
{
	var pheight = window.screen.availHeight;
	var pwidth = window.screen.availWidth;
	var pTop = (pheight - 656) / 2;
	var pLeft = (pwidth - 760) / 2;

	window.open("../ezEmail/mail_write_new_private.asp?type__=<%=type__%>", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 720px, width = 820px, status = no, toolbar=no, menubar=no,location=no,resizable=1");

}

	function AutoCalc(i_num){
		var i;
		//alert("<%=DbRec.Recordcount%>");

		//alert("<%=ipp%>");
		if("<%=DbRec.Recordcount%>" != 1){
			document.frmOutbox['Rec_Img'][i_num].src = "../ezPortal/Home/images/New_empty.gif"
		} else {
			document.frmOutbox['Rec_Img'].src = "../ezPortal/Home/images/New_empty.gif"
		}

	}






</script>
<script language=vbscript>
	function GetTimeCalcu(pDate , AddHour , AddMinute)
		pDate = dateadd("n" , AddMinute ,pDate)
		pDate= dateadd("h" , AddHour , pDate)
		if(Len(FormatDateTime(pDate,2)) >9) then
		        GetTimeCalcu = FormatDateTime(pDate, 2) &"T"& FormatDateTime(pDate, 4)
		else
		        GetTimeCalcu = "20" &  FormatDateTime(pDate, 2) &"T"& FormatDateTime(pDate, 4)
		end if 
	end function 
</script>
</HEAD>
<body style="BEHAVIOR:url('#default#userData');OVERFLOW:hidden" id="theBody" class="mainbody">


<% 
if type__ ="" then
	aaaa = "대표이사->발주청"
elseif type__ = "1" then
	aaaa = "책임감리원->감리본부"
elseif type__ = "2" then
	aaaa = "유관기관,기타"
else
	aaaa = ""
end if

%>


<table class="layout">
  <tr>


	<% 'if db_level = "P" OR db_level = "Z" THEN %>
		<td valign="top" height="40"><h1>요청자료접수함</h1>
	<% 'else %>
		<!--td valign="top" height="40"><h1>공문접수함</h1-->
	<% 'end if %>


      <div class="page">

		<img src="../../Home/images/page_previous.gif" width="15" height="15" align="absmiddle" hspace="2" id="td_Previous" onClick="goToPage('front')">

		페이지: <span id="td_pTotalCount"></span> <%=Nanum%> &nbsp;의
				
				<input name="Cnum" type="text" id="txt_PageInputNum" onkeypress="javascript : if (event.keyCode == 13) goToPage('page');" onselectstart="event.cancelBubble=true;event.returnValue=true" value="<%=curpage%>">

		        <img src="../../Home/images/page_next.gif" width="15" height="15" align="absmiddle" hspace="2" id="td_Previous" onClick="goToPage('next')">

	  </div>

	  	<div id="mainmenu">
        <ul id="tb_Parent">

			<% if db_level = "Z" THEN  %>
				<% if type__ <> ""  then %>
					<!--li><span onClick="Del_St(<%=curpage%>);">삭제</span></li-->
				<% END IF %>
			<% END IF %>


			<% if db_id = "204112" then %>

					<!--li><span onClick="ItemRead_onclick_test()"><img src="../../Home/images/i_mail.gif" alt=""  border="0" width="13" height="9">test</span></li-->

			<% end if %>
		  
		  <% if type__ <> "" THEN %>
			  <% 'if db_level = "P" OR db_level = "Z" THEN 
				'if db_level = "Z" THEN  %>
					<!--li><span onClick="new_mail_onclick1()"><img src="../../Home/images/i_mail.gif" alt=""  border="0" width="13" height="9">수기 작성</span></li-->
			  <% 'END IF %>
		  <% END IF %>		  

		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <li style="background:none;padding:0">          
            <select name="select" onChange="on_changeView(select.value)" style="WIDTH:110px">       
              <option VALUE="제목" selected>제목</option>
              <option VALUE="보낸이">보낸이</option>
              <option VALUE="현장명">현장명</option>
              <option VALUE="읽지않은공문">읽지않은공문</option>
            </select>
          </li>
          
		   <input type='textbox' size ="15" name='Gubun_i' VALUE="<%=Qgubun%>" onkeypress="javascript : if (event.keyCode == 13) goToPage1('<%=curpage%>');">
		  
           <li><span onClick="goToPage1(<%=curpage%>);" >검색</span></li>

        </ul>
      </div>
	  












<table class="mainlist" id ='test'>

  <form name="frmOutbox" action="BoardItemList.aspx" method="post">    
	<tr>

      <Th width=20 >
		<input type='checkbox' name="cbox" onclick='initCheckBehavior()' >
	  </Th>

      <!--th width=50 >번호</th-->
      <th width=20 ></th>

      
      <th style="cursor:hand;" width="200px" >제목</th>
      
      <th style="cursor:hand;" width="70px" >
		<% if type__ = "2" THEN %>
			보낸이
		<% else %>
			보낸이
		<% end if %>		
	  </th>
      
      <th style="cursor:hand;" width="100px" >
	  
		<% if type__ = "2" THEN %>
			문서번호
		<% elseif type__ = "1" THEN %>

			  <% if type__ <> "" THEN %>
				  <% 'if db_level = "P" OR db_level = "Z" THEN 
					if db_level = "Z" THEN  %>
						
				  <% else %>
					  현장명					
				  <% END IF %>
			  <% END IF %>	

		<% else %>
			현장명
		<% end if %>			  
	  
	  </th>
      
      <th style="cursor:hand;" width="70px" >받은날</th>

      <th style="cursor:hand;" width="40px" >상태</th>
      
      <th style="cursor:hand;padding:0" align="center" width="30px" ><img src="../../Home/images/file.gif" width="13" height="12"></th>
      
	  <th style="cursor:hand;"  width="30px" >
	  <% if UCASE(db_level) = "P" OR UCASE(db_level) = "Z" THEN %>
	      승인
	  <% END IF %>
	  </th>

      <th style="cursor:hand;"  width="50px" >
			전송여부	  
	  </th>


    </tr>




<% if postcount <> 0 then 

	qq = 0
	
	'if (curpage-1) * ipp > DbRec.Recordcount then

	'else
%>


<%


d = 1

 if d = 1 then



	For i = 1 to ipp
'	For i = 1 to cnt_q
		if totpage = curpage then
			value = postcount Mod ipp
			if i > value and value <> 0 then
				Exit For
			end if
		end if
%>
		<% visited = DbRec("o_visited")
		  if visited = 0 then
				sState = "안읽음"          
		  else
				sState = "읽음"		  
		  end if%>

		<%'if len(DbRec("o_subject")) > 30 then%>

		
		<%'else%> 
		<%'end if%> 



		<%
			send_date = DbRec("o_send_date")
			send_date = convertDate(send_date)


								file = RTRIM(LTRIM(DbRec("o_filename")))

								If file <> "" Then

								else

								end if
								

								qr = "Rd" & qq

								Qw = "Rq" & qq

								Aw = "Rq" & qq

								type1__ = DbRec("private__")
								
							%>



							<TR>
								
								<TD >
									<!--input type='checkbox' name='<%=qr%>' id='chk'-->
									<input type='checkbox' name='ccBox' id='chk' value="<%=DbRec("o_seq")%>" >
									<input type="hidden" name="db_acc" 		value="<%=DbRec("o_send_id")%>">
									<input type="hidden" name="type_var" 		value="<%=DbRec("private__")%>">
									<input type="hidden" name="O_Del" 		value="<%=DbRec("number_")%>">
								</td>

								<!--TD style="cursor:hand;"><%=DbRec("o_seq")%></td-->

								<TD style="cursor:hand;">
									<% IF DbRec("o_visited") = 0 THEN %>
										<IMG name="Rec_Img" SRC="../ezPortal/Home/images/New.gif" border="0">
									<% else %>
										<IMG name="Rec_Img" SRC="../ezPortal/Home/images/New_empty.gif" border="0">
									<% end if %>
								</td>

<TD title='' style='cursor:hand;text-overflow:ellipsis; overflow:hidden;' 								                                                    onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_receive_id")%>","접수","<%=DbRec("o_send_longname")%>","<%=visited%>","<%=Aw%>","<%=qq%>","<%=DbRec("private__")%>")'><nobr><%=DbRec("o_subject")%></nobr>
								</TD>

<TD style="cursor:hand;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_receive_id")%>","접수","<%=DbRec("o_send_longname")%>","<%=visited%>","<%=Aw%>","<%=qq%>","<%=DbRec("private__")%>")'>		
									<% if type1__ = "1" or type1__ = "2" THEN %>
										<%=DbRec("private_name")%>
									<% else %>
										<%=DbRec("o_send_name")%>
									<% end if %>									
								</TD>

<TD style="cursor:hand;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_receive_id")%>","접수","<%=DbRec("o_send_longname")%>","<%=visited%>","<%=Aw%>","<%=qq%>","<%=DbRec("private__")%>")'>
									<% if type1__ = "1" or type1__ = "2" THEN %>
										<%=DbRec("private_number")%>
									<% else %>
										<%=DbRec("o_send_longname")%>
									<% end if %>																		
								</TD>

<TD style="cursor:hand;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_receive_id")%>","접수","<%=DbRec("o_send_longname")%>","<%=visited%>","<%=Aw%>","<%=qq%>","<%=DbRec("private__")%>")'><%=send_date%> 
								</TD>

<TD name="State" id="State" style="cursor:hand;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_receive_id")%>","접수","<%=DbRec("o_send_longname")%>","<%=visited%>","<%=Aw%>","<%=qq%>","<%=DbRec("private__")%>")'>
								
								<div id='<%=Aw%>' ><%=sState%></div>
								
								</TD>

<TD style="cursor:hand;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_receive_id")%>","접수","<%=DbRec("o_send_longname")%>","<%=visited%>","<%=Aw%>","<%=qq%>","<%=DbRec("private__")%>")'>

									<%

										Set DbRec2=Server.CreateObject("ADODB.Recordset")
										DbRec2.CursorType=1
										sqlstr 	= "select o_seq, o_savefile, o_savepath from save_file where o_seq = " & DbRec("o_seq") & " "

										DbRec2.Open sqlstr, DbCon
										

										if DbRec2.Recordcount <> 0 then %>
									
											<img src="../../Home/images/file.gif" width="13" height="12">
									<%	ELSE %>


									<%	end if
										Set DbRec2=NOTHING
									%>
								</TD>

								<TD >
									<% if UCASE(db_level) = "P" OR UCASE(db_level) = "Z" THEN %>
										<% IF DbRec("number_") = "0" THEN %>
											<!--input type='checkbox' name='<%=Qw%>' id='chk' onclick='checkBox_checked("<%=DbRec("o_seq")%>")'-->
											<input type='checkbox' name='cc_Box' id='chk' onclick='checkBox_checked("<%=DbRec("o_seq")%>",this,"<%=curpage%>")'>
										<% ELSE %>
											<input type='checkbox' name='cc_Box' id='chk' onclick='checkBox_checked("<%=DbRec("o_seq")%>",this,"<%=curpage%>")' checked=TRUE >

										<% END IF %>
									<% END IF %>
								</TD>

								<TD >

									<% if DbRec("savehtml_") = "" THEN %>
											미전송
									<% ELSE %>
											전송
									<% END IF %>

								</TD>
								


							</TR>

		
	<%						qq = qq + 1
	DbRec.MovePrevious
	Next
	%>


<% 

end if
end if %>


<!--input type='textbox' size ="200" value='<%=str%>'-->
  </form>
</table>


    </td>
  </tr>
  <!--tr>
    <td><div id="idMsgViewer" style="BEHAVIOR:url(../ezEmail/Controls/view.htc);OVERFLOW:auto;width:100%;HEIGHT:100%" onPageChange="updateContext()" onRefreshPage="updateContext()" onSelectItem="prevShow()" acceptLang="ko" setTimezone="" rowsPerPage="10"></div></td>
  </tr>
  <tr id="tb_PrevShow" onMouseMove="move_preViewWindow()" onMouseDown="down_preViewWindow()" onMouseUp="up_preViewWindow()" style="DISPLAY:none; WIDTH:100%; HEIGHT:100px">
    <td>
	
		<table  border="0" cellspacing="0" cellpadding="0" style="border:1px solid #B5B5B5;OVERFLOW:hidden; CURSOR:move;" bgcolor="e4e4e4"  id="title_preview" onselectstart="event.cancelBubble = true, event.returnValue = false;" width="100%" class="viewtxt">
			<tr>
			  <td height="16" nowrap id="td_SndName" style="padding:2px 5px">보낸사람 :</td>
			  <td id="value_1" width="43%"><div id="div_SndName" style="OVERFLOW:hidden">&nbsp;</div></td>
			  <td nowrap id="td_Ref" style="padding:2px 5px">참&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;조 :</td>
			  <td id="td_divref" width="43%"   ><div style="OVERFLOW-Y: hidden; OVERFLOW-X: hidden; PADDING-TOP: 1px; HEIGHT: 15px" id="div_Ref" valign='center'></div></td>
			</tr>
			<tr>
			  <td height="16" valign="top" nowrap id="td_RcvName" style="padding:2px 5px">받는사람 :</td>
			  <td id="value_2" valign="top"><div style="OVERFLOW-Y:hidden; OVERFLOW-X:hidden; HEIGHT:14px" id="div_RcvName"></div></td>
			  <td valign="top" nowrap id="td_Attachment" style="padding:2px 5px">파일첨부 :</td>
			  <td valign="top" style="OVERFLOW: hidden"><span style="HEIGHT: 20px;overflow-y:auto;width:98%" id="div_Attachment" onMouseDown="event.cancelBubble=true">&nbsp;</span></td>
			</tr>
			<tr>
			  <td height="16" valign="top" nowrap id="td_Subject" style="padding:2px 5px">제&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;목 :</td>
			  <td style="OVERFLOW: auto" colspan="3" valign="top" style="padding:2px"><div style="OVERFLOW: hidden; HEIGHT:17px" id="div_Subject"></div></td>
			</tr>
		  </table>
      <div style="OVERFLOW:auto; WIDTH:100%; HEIGHT:100%; padding-top:5px" id="div_PreView" onselectstart="event.cancelBubble=true;event.returnValue=true"></div></td>
  </tr-->
</table>
<!--  받은편지함 모두삭제2008.01.14 이성조 -->
<!--table class="content" style="display:none">
  <tr>
    <td class="pos1">
	<div style="behavior:url(Controls/treeview.htc);border:0px solid B6B6B6;height:270;width:100%;overflow-x:auto;overflow-y:auto;background-color:#FFFFFF;padding-left:4px" id="PostTreeView" onnodeselect="PostTreeView.toggle(PostTreeView.selectedIndex)" onrequestdata="requestdata()">
	</div></td>
  </tr>
</table-->
<!--input type='textbox' size ="200" value='<%=str%>'-->
<!-- 끝. -->

<% set rs = Nothing %>
<% set DbRec = Nothing %>
<% set DbRec_Mis = Nothing %>



</body>
</HTML>



