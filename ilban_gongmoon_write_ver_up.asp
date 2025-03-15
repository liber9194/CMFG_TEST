<%@ LANGUAGE="VBSCRIPT" %>

<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=11">	

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
	strSql = strSql &   " ,'ilban_gongmoon_write_ver_up.asp' "
	strSql = strSql &   " ,'요청자료발송함 목록 [type=" & Request("type__") & "]' "
	strSql = strSql &   " ) "

	Set Result = DbCon.execute(strSql)
	Set Result=Nothing
	
'페이지 접속로그 추가 2016.04.21==================================================

number		= Request("number")

QSelect  = Request("QSelect")

Qgubun   = Request("Qgubun")

type__  = Request("type__")

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

db_id 	 	= session("db_id")
db_level 	 	= session("db_level")


Set DbRec=Server.CreateObject("ADODB.Recordset")
DbRec.CursorType=1

Set DbRec2=Server.CreateObject("ADODB.Recordset")
DbRec2.CursorType=1



	str = "SELECT A.o_seq, A.o_send_id, A.o_send_name, A.o_receive_id,  A.o_subject, A.o_content, A.o_filename, "
	str = str & " A.o_filesize, A.o_doc_no, A.o_send_date, A.o_send_longname,A.office_del, "
	str = str & " ISNULL(C.number_,'') AS number_, ISNULL(C.year_,'') AS year_, ISNULL(C.sabun,'') AS sabun, ISNULL(C.result_date,'') AS result_date  "
	str = str & " from office_tbl_i A "
	str = str & " LEFT join result_tbl_i C ON A.o_seq = C.o_seq "

	'개인별로 하는

	if db_level = "P" OR db_level = "Z" THEN
		str = str & " where A.o_send_id = '" & db_id & "' "
	ELSE
		'str = str & " where A.o_send_id = '" & site_code & "' "
		str = str & " where (A.o_send_id = '" & site_code & "' or A.o_send_id = '" & db_id & "') "
	end if


	str = str & " and A.type__ = '" & type__ & "' and A.private__ = '' "
	str = str & " and A.office_del = '' "


	'str = str & " where A.o_send_id = '" & db_id & "' "



	str = str & " and A.office_del = '' "

	if QSelect = "제목" then
		str = str & " and A.o_subject Like '%" & Qgubun & "%' "
	end if
	if QSelect = "읽지않은공문" then
		'str = str & " and B.o_visited < 1 "
	end if

	if QSelect = "보낸날" THEN
		str = str & " and A.o_send_date like '" & Qgubun & "%' "
	END IF
	
	IF QSelect = "현장명" then
		str = str & " and A.o_send_longname Like '" & Qgubun & "%' "
	END IF




	str = str & " order by A.o_seq asc "



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


		sql="SELECT Count(*) as totalcount from office_tbl_i "
		sql = sql & " where (o_send_id = '" & db_id & "' ) and (o_sdel_flag = 0)"


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
<script language="JScript" src="../ezEmail/lang/ezEmail_ko.js"></script>
<script language="JavaScript" src="../../Home/myoffice/common/mouseeffect.js"></script>
<STYLE> 
P { MARGIN-BOTTOM: 0mm; MARGIN-TOP: 0mm } 
</STYLE>

<!--<script language="JScript" src="../ezEmail/js/emails.js"></script>-->
<script language="JScript" src="../ezEmail/js/email_tree.js"></script>
<script language="JScript" src="../ezEmail/js/string_component.js"></script>
<script>
 
var end_page = "<%=Nanum%>";

	/*
	function window.onload() 
	{
		switch (g_foldertype)
		{
			case "sent":
				receivecheck.style.display='';
				reply.style.display='none';
				select.selectedIndex = 5; //보낸 편지함이면 셀렉트 박스를 받은사람 정렬로 변경한다.
				break;
			case "draft":
				g_bdraft=true;
				break;
			case "delete":
				deleteone.style.display='none';
				deleteall.style.display='';
				break;
		}
		
		g_ViewID = idMsgViewer;
		g_PageCount = td_pTotalCount;
		g_PageInput = txt_PageInputNum;

		g_PreViewID = tb_PrevShow;
		g_PreView = div_PreView;
		g_PreviewTitle = title_preview;
		GetInfo();
		
		window.setInterval(getUnReadCount, 1000 * 300);
		preViewSizeSetting();

		theBody.load("valueStore");
		if (theBody.getAttribute("preView") != "OFF") 
			prevShow_onclick();

		window.onresize();
		window.focus();
		if( g_foldertype != "sent" && g_foldertype != "draft" )
			btnReject.style.display = "";
		
		
		
        //-----------받은편지함 모두삭제2008.01.14 이성조-----------//
		window.returnValue = 0;
		var xmlDom = new ActiveXObject("Microsoft.XMLDom");
		xmlDom.async = false;
		xmlDom.load("Controls/tree_config.xml");
		PostTreeView.config = xmlDom;
		
		if( g_ExchangeVS == "2007" )
		{
			PostTreeView.source = "<tree><nodes>" + get_childXML("http://EXMAIL/exchange/204112/", true, false) +
					"</nodes></tree>";
					
		}
		else
		{
			PostTreeView.source = "<tree><nodes>" + get_childXML("http://gw.dohwa.co.kr/exchange/204112/", true, false) +
					"</nodes></tree>";
		}
		
		PostTreeView.update();
		xmlDom = null;
        //--- 끝. ---//				
        
	}
    function sleep(sec) 
    {
        var now = new Date();
        var exitTime = now.getTime() + (sec*1000);
        while (true) {
            now = new Date();
            if (now.getTime() > exitTime) return;
        }
    }
	function Received_MailALLD()
	{
        if (confirm("편지함에 있는 메일을 모두 삭제하시겠습니까?"))
        {
            var deleteURL = PostTreeView.getvalue(4, "href");
            showProgress("받은 편지함을 전체 삭제 진행중 입니다");
		    var result = delete_mail(PostTreeView.getvalue(1, "href"), false, deleteURL);
		    if (result == 100){hideProgress();
			    alert("삭제할 메일이 없습니다.");}
		    else if (result != true){hideProgress();
			    alert("메일 삭제중 에러발생.");}
		    else{hideProgress();
			    alert("메일을 모두 삭제하였습니다.");}
				
		    refreshUnreadCount();
		    refresh_onclick();
		}
	}

	function window.onunload()
	{
		if (g_bPrevShow == true) 
			theBody.setAttribute("preView", "ON");
		else 
			theBody.setAttribute("preView", "OFF");
		theBody.save("valueStore");
	} 

	function window.onresize()
	{
	}
	*/

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
		if (confirm("요청자료를 삭제하시겠습니까?")) {      
			var oColl = document.frmOutbox.ccBox;
			var str = '';			
			
			if ("<%=DbRec.Recordcount%>" == 1){
						
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

					parent.frames[2].location  = "GongMoonDel_Office_i_ver_up.asp?Page=" + Cpage + "&Str=" + str + "&type__=<%=type__%>";
				}

			} else {
					
				for (var i=0; i < oColl.length; i++) {
																
					if (oColl[i].checked==true){
						
						if (document.frmOutbox.O_Del[i].value > 0) {
							alert("공문번호 가 발급된 공문은 삭제 할수 없습니다");
							str = '';
							break;
						}
						str += oColl[i].value + ';';
					}

				}
				
				if (str == ''){
				}else {

					parent.frames[2].location  = "GongMoonDel_Office_i_ver_up.asp?Page=" + Cpage + "&Str=" + str + "&type__=<%=type__%>";
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

		window.location.href = "ilban_gongmoon_write_ver_up.asp?Page=" + aaa1 + "&type__=<%=type__%>" ;
	}


	function goToPage1(aaa1)
	{
		window.location.href = "ilban_gongmoon_write_ver_up.asp?Page=" + aaa1 + "&QSelect=" + document.all.select.value + "&Qgubun=" + document.all.Gubun_i.value + "&type__=<%=type__%>"; 
	}

	/*
	function document.onselectstart()
	{
		event.cancelBubble = true;
		event.returnValue = false;
	}
	*/

	function SortPage(SortBy)
	{
		window.location.href = "../ezEmail/mail_read_i_ver_up.asp?Seq=" + SortBy ;
	}

	function ItemRead_onclick(pItemBoardID,rID,Stype,HJname)
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
		
		window.open("../ezEmail/mail_read_i_ver_up.asp?Seq=" + pItemBoardID + "&rID=" + rID + "&stype=" + Stype + "&HJname=" + HJname , "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,height=700,width=765,top=" + pTop + ",left=" + pLeft, "");	
	}
	
	function new_mail_onclick1() 
	{
		if ("<%=session("db_id")%>" == "")
		{
			alert("세션이 만료되었습니다.\n\n재로그인 후 이용하여 주십시오.");
			return;
		}

		var pheight = window.screen.availHeight;
		var pwidth = window.screen.availWidth;
		var pTop = (pheight - 656) / 2;
		var pLeft = (pwidth - 760) / 2;		

		window.open("../ezEmail/mail_write_new_i_ver_up.asp?type__=<%=type__%>", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 720px, width = 820px, status = no, toolbar=no, menubar=no,location=no,resizable=1");
	}
	
	function new_mail_onclick1_test() 
	{
		if ("<%=session("db_id")%>" == "")
		{
			alert("세션이 만료되었습니다.\n\n재로그인 후 이용하여 주십시오.");
			return;
		}

		var pheight = window.screen.availHeight;
		var pwidth = window.screen.availWidth;
		var pTop = (pheight - 656) / 2;
		var pLeft = (pwidth - 760) / 2;		

		window.open("../ezEmail/mail_write_new_i_test2.asp?type__=<%=type__%>", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 720px, width = 820px, status = no, toolbar=no, menubar=no,location=no,resizable=1");
	}
	
	/*	
	function new_mail_onclick1_test2() 
	{
		var pheight = window.screen.availHeight;
		var pwidth = window.screen.availWidth;
		var pTop = (pheight - 656) / 2;
		var pLeft = (pwidth - 760) / 2;			

		//alert("mail_write_new_i");

		window.open("../ezEmail/mail_write_new_i_test2_JH.asp?type__=<%=type__%>", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 720px, width = 820px, status = no, toolbar=no, menubar=no,location=no,resizable=1");
	}
	*/
	
	function getReceiverList()
	{
		var pheight = window.screen.availHeight;
		var pwidth = window.screen.availWidth;
		var pTop = (pheight - 850) / 2;
		var pLeft = (pwidth - 800) / 2;		

		window.open("../Addr_New/select_requestBoard_chk_ver_up.asp", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 785px, width = 1500px, status = no, toolbar=no, menubar=no,location=no,resizable=1");
	}

	function go_alert()
	{
		var pheight = window.screen.availHeight;
		var pwidth = window.screen.availWidth;
		var pTop = (pheight - 850) / 2;
		var pLeft = (pwidth - 800) / 2;		

		window.open("go_alert_test.asp", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 730px, width = 1500px, status = no, toolbar=no, menubar=no,location=no,resizable=1");
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

<table class="layout" style="width: 100%;">
	<tr>
	    <td valign="top" height="40">
		
			<h1>요청자료발송함</h1>

			<div class="page">
				<img src="../../Home/images/page_previous.gif" width="15" height="15" align="absmiddle" hspace="2" id="td_Previous" onClick="goToPage('front')">

				페이지: <span id="td_pTotalCount"></span> <%=Nanum%> &nbsp;의
					<input name="Cnum" type="text" id="txt_PageInputNum" onkeypress="javascript : if (event.keyCode == 13) goToPage('page');" onselectstart="event.cancelBubble=true;event.returnValue=true" value="<%=curpage%>">
					<img src="../../Home/images/page_next.gif" width="15" height="15" align="absmiddle" hspace="2" id="td_Previous" onClick="goToPage('next')">
			</div>

			<div class="row" style="padding-top: 43px;padding-bottom:5px">
				<div class="col-sm-6" style="text-align:left;">
					<div class="form-group form-group-jh">
						<div class="col-sm-5" style="padding-left:0px;padding-right:0px;">
							<button type="button" class="btn btn-default btn-jh" onClick="new_mail_onclick1();">요청자료 쓰기</button>
							<button type="button" class="btn btn-default btn-jh" onClick="Del_St(<%=curpage%>);">삭제</button>
							
							<% if db_level = "Z" then %>
								<button type="button" class="btn btn-default btn-jh" onClick="getReceiverList();">요청자료 조회자 확인</button>
							<% end if %>							
						</div>
						<div class="col-sm-2" style="padding-left:5px;padding-right:0px;">
						
							<select class="form-control input-sm" name="select" onChange="on_changeView(select.value)">       
								<option VALUE="제목" selected>제목</option>				
							</select>
						
						</div>
						<div class="col-sm-3" style="padding-left:5px;padding-right:0px;">
							<input class='form-control' type='textbox' size ="15" name='Gubun_i' VALUE="<%=Qgubun%>" onkeypress="javascript : if (event.keyCode == 13) goToPage1('<%=curpage%>');">
						</div>
						<div class="col-sm-2" style="padding-left:5px;">
							<button type="button" class="btn btn-default btn-jh" onClick="goToPage1(<%=curpage%>);">검색</button>
						</div>
					</div>
				</div>
				<div class="col-sm-6" style="text-align:left;">
					<% if db_id = "216050" then %>						
						<button type="button" class="btn btn-default btn-jh" onClick="go_alert();">uc메신저 test</button>
					<% end if %>
				</div>
			</div>	
			
			<table class="mainlist" id ='test'>
				<form name="frmOutbox" action="BoardItemList.aspx" method="post">    
				<tr>

					<Th width=20 style="text-align:center;">
						<input type='checkbox' name="cbox" onclick='initCheckBehavior()' >
					</Th>

					<!--th width=50 >번호</th-->
		  
					<th style="cursor:pointer;" width="230px" >제목</th>
		  
					<th style="cursor:pointer;" width="80px" >보낸이</th>
		  
					<th style="cursor:pointer;" width="100px" >현장명</th>
		  
					<th style="cursor:pointer;text-align:center;" width="80px" >보낸날</th>

					<!--th style="cursor:pointer;" width="30px" >상태</th-->
		  
					<th style="cursor:pointer;padding:0;text-align:center;" align="center" width="30px" ><img src="../../Home/images/file.gif" width="13" height="12"></th>
		  
					<th style="cursor:pointer;"  width="20px" >
					</th>
				</tr>

				<% if postcount <> 0 then 

						qq = 0

						For i = 1 to ipp
							if totpage = curpage then
								value = postcount Mod ipp
								if i > value and value <> 0 then
									Exit For
								end if
							end if

							send_date = DbRec("o_send_date")
							send_date = convertDate(send_date)

							file = RTRIM(LTRIM(DbRec("o_filename")))							

							qr = "Rd" & qq
							Qw = "Rq" & qq
							Aw = "Rq" & qq						
				%>

							<TR>							
								<TD >								
									<input type='checkbox' name='ccBox' id='chk' value="<%=DbRec("o_seq")%>" >
									<input type="hidden" name="db_acc" 		value="<%=DbRec("o_send_id")%>">
									<input type="hidden" name="O_Del" 		value="<%=DbRec("number_")%>">
								</td>

								<!--TD style="cursor:pointer;"><%=DbRec("o_seq")%></td-->

								<TD title='' style='cursor:pointer;text-overflow:ellipsis; overflow:hidden;text-align:left;' onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_send_id")%>","발송","<%=DbRec("o_send_longname")%>")'><nobr><%=DbRec("o_subject")%></nobr>
								</TD>
									
								<TD style="cursor:pointer;text-align:left;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_send_id")%>","발송","<%=DbRec("o_send_longname")%>")'><%=DbRec("o_send_name")%></TD>

								<TD style="cursor:pointer;text-align:left;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_send_id")%>","발송","<%=DbRec("o_send_longname")%>")'><%=DbRec("o_send_longname")%></TD>

								<TD style="cursor:pointer;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_send_id")%>","발송","<%=DbRec("o_send_longname")%>")'><%=send_date%> 
								</TD>							

								<TD style="cursor:pointer;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_send_id")%>","발송","<%=DbRec("o_send_longname")%>")'>
									<%
										Set DbRec2=Server.CreateObject("ADODB.Recordset")
										DbRec2.CursorType=1
										sqlstr 	= "select o_seq, o_savefile, o_savepath from save_file_i where o_seq = " & DbRec("o_seq") & " "

										DbRec2.Open sqlstr, DbCon									

										if DbRec2.Recordcount <> 0 then %>								
											<img src="../../Home/images/file.gif" width="13" height="12">
									<%	ELSE %>

									<%	end if
										Set DbRec2=NOTHING
									%>
								</TD>

								<TD>								
								</TD>
							</TR>		
		<%					qq = qq + 1

							DbRec.MovePrevious
						Next
					end if 
		%>
				</form>
			</table>
		</td>
	</tr> 
</table>
<% set DbRec = Nothing %>
<% set DbRec2 = Nothing %>
<% set Result = Nothing %>
<% set Result1 = Nothing %>
<% set rs = Nothing %>
</body>
</HTML>