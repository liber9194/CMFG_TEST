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
	strSql = strSql &   " ,'ilban_gongmoon_manager.asp' "
	strSql = strSql &   " ,'요청자료관리함 목록' "
	strSql = strSql &   " ) "

	Set Result = DbCon.execute(strSql)
	Set Result=Nothing
	
'페이지 접속로그 추가 2016.04.21==================================================

number		= Request("number")



QSelect  = Request("QSelect")

Qgubun   = Request("Qgubun")


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


S_Gubun = Request("S_Gubun")

if trim(S_Gubun) <> "" then
	curpage=1
end if

'if trim(Qgubun) <> "" then
'	curpage=1
'end if




ipp=15
ten=5

db_id 	 	= session("db_id")
db_level 	 	= session("db_level")


Set DbRec=Server.CreateObject("ADODB.Recordset")
DbRec.CursorType=1

Set DbRec2=Server.CreateObject("ADODB.Recordset")
DbRec2.CursorType=1



	str = "SELECT A.o_seq, A.o_send_id, A.o_send_name, A.o_receive_id,  A.o_subject, A.o_content, A.o_filename, "
	str = str & " A.o_filesize, A.o_doc_no, A.o_send_date, A.o_send_longname,A.file_Result, "
	str = str & " ISNULL(C.number_,'') AS number_, ISNULL(C.year_,'') AS year_, ISNULL(C.sabun,'') AS sabun, ISNULL(C.result_date,'') AS result_date, A.type__,[private__],private_number,private_name  "
	'str = str & " ISNULL(D.s_id,'') AS S_II "
	str = str & " from office_tbl_i A "
	'str = str & " LEFT join RECEIVE_TBL_i B ON A.o_seq = B.o_seq AND  B.o_visited > 0 "
	str = str & " LEFT join result_tbl_i C ON A.o_seq = C.o_seq "
	'str = str & " LEFT join office_Manager D ON A.o_seq = D.seq AND D.s_id = '" & db_id & "' "
	str = str & " Where A.o_send_date >= Convert(Varchar(10), DateAdd(mm, -1, GetDate()), 120) "
	str = str & " and A.o_seq not in ( select g.seq from office_Manager_i g where g.s_id = '" & db_id & "' and g.Del_Chk='1') "
'	str = str & " where A.o_send_id = '" & db_id & "' "
'	str = str & " and A.office_del = '' "

	if QSelect = "제목" then
		str = str & " and A.o_subject Like '%" & Qgubun & "%' "
	end if
	if QSelect = "읽지않은공문" then
		'str = str & " and B.o_visited < 1 "
	end if

	if QSelect = "보낸이" THEN
		str = str & " and A.o_send_name Like '" & Qgubun & "%' "
	END IF
	
	IF QSelect = "현장명" then
		str = str & " and A.o_send_longname Like '" & Qgubun & "%' "
	END IF

	IF QSelect = "받는사람" then
		str = str & " and A.o_seq in (select h.o_seq  from receive_tbl_i h where h.receive_del = '' and replace(h.o_receive_name,' ','') like '" & Qgubun & "%') "
	END IF

	IF QSelect = "공문번호(대,발)" then
		str = str & " and A.type__ = '' and  number_  = " & Qgubun & " "
	END IF

	IF QSelect = "공문번호(책,감)" then
		str = str & " and A.type__ = '1' and  number_  = " & Qgubun & " "
	END IF

	IF QSelect = "공문번호(유,기)" then
		str = str & " and A.type__ = '2' and  number_  = " & Qgubun & " "
	END IF


             ' <option VALUE="받는사람">받는사람</option>
			 ' <option VALUE="공문번호(대,발)">공문번호(대,발)</option>
			 ' <option VALUE="공문번호(책,감)">공문번호(책,감)</option>
			 ' <option VALUE="공문번호(유,기)">공문번호(유,기)</option>


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
	var g_ExchangeVS = "2007";
	var g_bdraft = false;
	var g_moveUrl = "http://exmail/exchange/204112/%EB%B0%9B%EC%9D%80%20%ED%8E%B8%EC%A7%80%ED%95%A8";
	var g_expath = "exchange";
	var g_userName = "이재훈";
	var g_szRootFolderName = '받은편지함';
	var g_exuserid = "204112";
	var g_author = "Basic RE9IV0EuQ08uS1JcMjA0MTEyOno4ODk1MjQ5";
	var g_bPrevShow = false;
	var g_ViewID = null;
	var g_PreViewID = null;
	var g_PageInput = null;
	var g_PageCount = 0;
	var g_PreView = null;
	var g_PreviewTitle = null;
	var g_moveStart = false;
	var g_startPosition = 0;
	var g_foldertype = "";
	var importanceColor = "BLUE";
	var g_userLang= "1";
	
	var g_timeset = "235|+09:00";
	
	var g_progresswin = null;	// 삭제진행중 화면 표기 2008.01.14 이성조 
	





var end_page = "<%=Nanum%>"




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









//window.onload = function()
//{
//  initCheckBehavior();
//}


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

//						var oColl = obj.form.elements;

					//alert(document.frmOutbox.ccBox.length);
						var oColl = document.frmOutbox.ccBox;

						for (var i=0; i < document.frmOutbox.ccBox.length; i++) {
							oColl[i].checked = document.frmOutbox.cbox.checked;
							//if (oColl[i].checked==1) total ++;
							//if (total >= 5){
							//	alert('5개 이상 체크했네요~');
							//	obj.checked = 0;
							//	break;
							//}
						}


	//alert(document.frmOutbox.Rd0.checked);


	//document.frmOutbox.Rd0.checked = document.frmOutbox.cbox.checked
	//document.frmOutbox.Rd1.checked = document.frmOutbox.cbox.checked
	//document.frmOutbox.Rd2.checked = document.frmOutbox.cbox.checked
	//document.frmOutbox.Rd3.checked = document.frmOutbox.cbox.checked
	//document.frmOutbox.Rd4.checked = document.frmOutbox.cbox.checked
	//document.frmOutbox.Rd5.checked = document.frmOutbox.cbox.checked
	//document.frmOutbox.Rd6.checked = document.frmOutbox.cbox.checked
	//document.frmOutbox.Rd7.checked = document.frmOutbox.cbox.checked
	//document.frmOutbox.Rd8.checked = document.frmOutbox.cbox.checked
	//document.frmOutbox.Rd9.checked = document.frmOutbox.cbox.checked
}










	
	//function document.onkeydown()
	//{

	//	if (window.event.keyCode == "37")
	//		goToPage("front");
	//	else if (window.event.keyCode == "39")			
	//		goToPage("next");
	//	else if (window.event.keyCode == "46")
	//	{
	//	aaa = 1;
			//if (event.shiftKey)
			//	deleteWork(true);
			//else
			//	deleteWork(false);
	//	}
	//}



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
				//document.all.Cnum.text

				//aaa = parseInt(document.all.Cnum.value,0) + 1

				if (end_page < aaa1) {
					aaa1 = end_page

				}

				if (1 > aaa1) {
					aaa1 = 1
				}


				//alert(aaa1);
			    //parent.frames[0].location = "right_main_gongmoon_all.asp?Page=" + aaa1 ;
				window.location.href = "ilban_gongmoon_manager.asp?Page=" + aaa1 + "&QSelect=<%=QSelect%>&Qgubun=<%=Qgubun%>"  ;
            //}


            //else
            //{
            //    window.open("BoardItemView.aspx?ShowAdjacent=" + ShowAdjacent + "&ItemID=" + pItemID + "&BoardID=" + pItemBoardID, "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,height=720,width=765,top=" + pTop + ",left=" + pLeft, "");
            //}
		}








		function goToPage1(aaa1)
		{


				//alert(document.all.select.value);
				//alert(document.all.Gubun_i.value);

				//alert(aaa1);
			//    parent.frames[0].location = "right_main_gongmoon_insert.asp?Page=" + aaa1 + "&QSelect=" + document.all.select.value + "&Qgubun=" + document.all.Gubun_i.value; 
				window.location.href = "ilban_gongmoon_manager.asp?Page=" + aaa1 + "&QSelect=" + document.all.select.value + "&Qgubun=" + document.all.Gubun_i.value + "&S_Gubun=1"; 
		}















	function document.onselectstart()
	{
		event.cancelBubble = true;
		event.returnValue = false;
	}

		function SortPage(SortBy)
		{
			window.location.href = "../ezEmail/mail_read.asp?Seq=" + SortBy ;
//			window.location.href = "../ezEmail/mail_read.asp?Seq=" + SortBy + "&BoardID=" + pBoardID + "&pBoardName=" + pBoardName + "&SortBy=" + SortBy;
		}






		function ItemRead_onclick(pItemBoardID,rID,Stype,HJname,qq_r)
		{
			//if(Read_FG != "true") {
			//	alert("읽기 권한이 없습니다.");
			//	return;
			//}


//alert("a");

			var e = event.srcElement;
			var eText = e.outerHTML;
			if(eText.substring(0,3)=="<B>"){
				e.outerHTML = eText.substring(3, eText.length);
			}
			
			var pheight = window.screen.availHeight;
			var pwidth = window.screen.availWidth;
			var pTop = (pheight - 720) / 2;
			var pLeft = (pwidth - 765) / 2;
			
			//if(gubun!="3")
			//{
			//alert(pItemBoardID);
			//alert(rID);
			//alert(Stype);
			//alert(HJname);
//alert("b");
			if ("<%=db_level%>" == "S") {
			
				AutoCalc(qq_r);
			}

//alert("c");

			    window.open("../ezEmail/mail_read_i.asp?Seq=" + pItemBoardID + "&rID=" + rID + "&stype=" + Stype + "&HJname=" + HJname + "&Cur=1&qip=1" , "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,height=700,width=765,top=" + pTop + ",left=" + pLeft, "");	
            //}
            //else
            //{
            //    window.open("BoardItemView.aspx?ShowAdjacent=" + ShowAdjacent + "&ItemID=" + pItemID + "&BoardID=" + pItemBoardID, "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,height=720,width=765,top=" + pTop + ",left=" + pLeft, "");
            //}
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













		function get_row(r,c) { 
		 alert(test.rows[r].cells[c].innerHTML);
		} 

		function ItemRead_onclick1(pItemBoardID,hjName,visited,V1)
		{


				//document.getElementById(V1).innerHTML="읽음";
			    //parent.frames[1].location = "right_main_gongmoon1_INSERT.asp?Seq=" + pItemBoardID + "&hjName=" + hjName + "&visited=" + visited;
				parent.frames[1].location = "right_main_gongmoon1_all.asp?Seq=" + pItemBoardID + "&hjName=" + hjName;
				
		}



		function checkBox_checked(pItemBoardID,obj,Cpage)
		{
			//if(Read_FG != "true") {
			//	alert("읽기 권한이 없습니다.");
			//	return;
			//}
			//var e = event.srcElement;
			//var eText = e.outerHTML;
			//if(eText.substring(0,3)=="<B>"){
			//	e.outerHTML = eText.substring(3, eText.length);
			//}
			
			//var pheight = window.screen.availHeight;
			//var pwidth = window.screen.availWidth;
			//var pTop = (pheight - 720) / 2;
			//var pLeft = (pwidth - 765) / 2;
			
			//if(gubun!="3")
			//{

				//alert("t");
			    //parent.frames[1].location = "right_main_gongmoon1.asp?Seq=" + pItemBoardID + "&hjName=" + hjName;
            //}


//Rq


						//alert(obj.checked);

			if (obj.checked == true) {






					//if document.frmOutbox.Rd0.checked = document.frmOutbox.cbox.checked


						if (confirm("공문번호 를 생성 하시겠습니까?")) {      
							// window.open("GongMoonNumberInsert.asp?Seq=" + pItemBoardID + "&Page=" + Cpage , "aaa", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,height=0,width=0,top=0,left=0", "");
							
							//window.location = "GongMoonNumberInsert.asp?Seq=" + pItemBoardID + "&Page=" + Cpage
							parent.frames[0].location  = "GongMoonNumberInsert.asp?Seq=" + pItemBoardID + "&Page=" + Cpage



								//if (form.FILE1.value==""){
								//}else{
								//	Bouncer();
								//}
								f_submit();
								//document.WebWrite_form.submit();

								return true;
						} else {

							obj.checked = 0;		
						}
			} else {

				
				obj.checked = 1;

			}

            //else
            //{
            //    window.open("BoardItemView.aspx?ShowAdjacent=" + ShowAdjacent + "&ItemID=" + pItemID + "&BoardID=" + pItemBoardID, "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,height=720,width=765,top=" + pTop + ",left=" + pLeft, "");
            //}
		}
















/*		function ItemRead_onclick(pItemBoardID, pItemBoardName, pItemID, pUserID)
		{
			if(Read_FG != "true") {
				alert("읽기 권한이 없습니다.");
				return;
			}
			var e = event.srcElement;
			var eText = e.outerHTML;
			if(eText.substring(0,3)=="<B>"){
				e.outerHTML = eText.substring(3, eText.length);
			}
			
			var pheight = window.screen.availHeight;
			var pwidth = window.screen.availWidth;
			var pTop = (pheight - 720) / 2;
			var pLeft = (pwidth - 765) / 2;
			
			if(gubun!="3")
			{
			    window.open("BoardItemView.aspx?ShowAdjacent=" + ShowAdjacent + "&ItemID=" + pItemID + "&BoardID=" + pItemBoardID, "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,height=720,width=765,top=" + pTop + ",left=" + pLeft, "");	
            }
            else
            {
                window.open("BoardItemView.aspx?ShowAdjacent=" + ShowAdjacent + "&ItemID=" + pItemID + "&BoardID=" + pItemBoardID, "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,height=720,width=765,top=" + pTop + ",left=" + pLeft, "");
            }
		}

*/



function new_mail_onclick1() 
{
	var pheight = window.screen.availHeight;
	var pwidth = window.screen.availWidth;
	var pTop = (pheight - 656) / 2;
	var pLeft = (pwidth - 760) / 2;

//	window.open("../ezEmail/mail_write.aspx?cmd=NEW", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 660px, width = 760px, status = no, toolbar=no, menubar=no,location=no,resizable=1");
	
	
	window.open("../ezEmail/mail_write.asp", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 660px, width = 760px, status = no, toolbar=no, menubar=no,location=no,resizable=1");
//	window.open("../Mis/WebWrite_asp/WebWrite.asp", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 660px, width = 760px, status = no, toolbar=no, menubar=no,location=no,resizable=1");



//	window.open("../Mis/WebWrite/action.asp", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 660px, width = 760px, status = no, toolbar=no, menubar=no,location=no,resizable=1");


}


function new_mail_onclick() 
{
	var pheight = window.screen.availHeight;
	var pwidth = window.screen.availWidth;
	var pTop = (pheight - 656) / 2;
	var pLeft = (pwidth - 760) / 2;

//	window.open("../ezEmail/mail_write.aspx?cmd=NEW", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 660px, width = 760px, status = no, toolbar=no, menubar=no,location=no,resizable=1");
	
	
//	window.open("../ezEmail/mail_write.asp", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 660px, width = 760px, status = no, toolbar=no, menubar=no,location=no,resizable=1");

	ALERT("T");
	window.open("../Mis/WebWrite_asp/WebWrite.asp", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 660px, width = 760px, status = no, toolbar=no, menubar=no,location=no,resizable=1");



//	window.open("../Mis/WebWrite/action.asp", "", "top=" + pTop.toString() + ", left=" + pLeft.toString() + ", height = 660px, width = 760px, status = no, toolbar=no, menubar=no,location=no,resizable=1");


}






//function Del_St(Cpage)
//{
//						if (confirm("공문 을 삭제 하시겠습니까?")) {      
//								var oColl = document.frmOutbox.ccBox;
//								var str = ''
//								for (var i=0; i < document.frmOutbox.ccBox.length; i++) {
//									if (oColl[i].checked==1){
//										str += oColl[i].value + ';';
//									}
//								}
//								if (str == ''){
//								}else {
//									parent.frames[0].location  = "GongMoonDel1.asp?Page=" + Cpage + "&Str=" + str;
//								}
//						} 
//}



function Del_St(Cpage)
{

						if (confirm("삭제 하시겠습니까?")) {      
								var oColl = document.frmOutbox.ccBox;
								var str = ''								
								
								if ("<%=DbRec.Recordcount%>" == 1){
											
											if (oColl.checked==true){												
													str += oColl.value + ';';
											}		
											
											if (str != ''){
												parent.frames[2].location  = "GongMoonDel_Manager_i.asp?Page=" + Cpage + "&Str=" + str + "&type__=<%=type__%>";
											}

								} else {
										
										for (var i=0; i < oColl.length; i++) {																						
											if (oColl[i].checked==true){
												str += oColl[i].value + ';';												
											}

										}
										
										if (str != ''){
											parent.frames[2].location  = "GongMoonDel_Manager_i.asp?Page=" + Cpage + "&Str=" + str + "&type__=<%=type__%>";
										}
								}
						} 


}


function Del_St1(Cpage)
{

						if (confirm("삭제 하시겠습니까?")) {      
								var oColl = document.frmOutbox.ccBox;
								var str = ''								
								
								if ("<%=DbRec.Recordcount%>" == 1){
											
											if (oColl.checked==true){												
													str += oColl.value + ';';
											}		
											
											if (str != ''){
												parent.frames[2].location  = "GongMoonDel_Manager_i_all.asp?Page=" + Cpage + "&Str=" + str + "&type__=<%=type__%>";
											}

								} else {
										
										for (var i=0; i < oColl.length; i++) {																						
											if (oColl[i].checked==true){
												str += oColl[i].value + ';';												
											}

										}
										
										if (str != ''){
											parent.frames[2].location  = "GongMoonDel_Manager_i_all.asp?Page=" + Cpage + "&Str=" + str + "&type__=<%=type__%>";
										}
								}
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



<table class="layout">
  <tr>
    <td valign="top" height="40"><h1>요청자료관리함</h1>



<!--input name="Cnu222m" type="text" id="txt_PageInputNum" onKeyDown="goToPage('page')" onselectstart="event.cancelBubble=true;event.returnValue=true" value="<%=STR%>"-->
      <div class="page">
		<!--a href="receive_board.asp?page=<%=prevpg%>&db_acc=<%=db_acc%>&code=<%=code%>&title=<%=title%>" 
					onmouseover="window.status=('이전 페이지로 가기');return true;"  
					onmouseout="window.status=('&nbsp;');return true;" id="RED"-->
		<img src="../../Home/images/page_previous.gif" width="15" height="15" align="absmiddle" hspace="2" id="td_Previous" onClick="goToPage('front')">
		<!--/a-->
		페이지: <span id="td_pTotalCount"></span> <%=Nanum%> &nbsp;의
				<input name="Cnum" type="text" id="txt_PageInputNum" onkeypress="javascript : if (event.keyCode == 13) goToPage('page');" onselectstart="event.cancelBubble=true;event.returnValue=true" value="<%=curpage%>">
				<!--a href="receive_board.asp?page=<%=nextpg%>&db_acc=<%=db_acc%>&code=<%=code%>&title=<%=title%>" 
					onmouseover="window.status=('다음 페이지로 가기');return true;"  
					onmouseout="window.status=('&nbsp;');return true;" id="RED"-->

		        <img src="../../Home/images/page_next.gif" width="15" height="15" align="absmiddle" hspace="2" id="td_Previous" onClick="goToPage('next')">
		<!--/a-->
	  </div>



		<div id="mainmenu">
        <ul id="tb_Parent">
          <!--li><span onClick="new_mail_onclick1()"><img src="../../Home/images/i_mail.gif" alt=""  border="0" width="13" height="9">공문쓰기</span></li>
          <li id="reply"><span onClick="new_mail_onclick()"><img src="../../Home/images/i_mailreply.gif" alt=""  border="0" width="13" height="9">회신</span></li>
          <li><span onClick="all_reply_mail_onclick()"><img src="../../Home/images/i_reall.gif" alt="" width="14" height="13"  border="0" align="absmiddle">전체회신</span></li>
          <li><span onClick="transmission_mail_onclick()"><img src="../../Home/images/i_fw.gif" alt=""  border="0" width="13" height="9">전달</span></li>
          <img src="../../Home/images/i_bar.gif" align="absmiddle">
          <li><span onClick="move_mail_onclick()">이동/복사</span></li-->
          <li><span onClick="Del_St(<%=curpage%>);">삭제</span></li>


			<% if UCASE(db_level) = "Z" THEN %>
					  <li><span onClick="Del_St1(<%=curpage%>);">임원관리 같이삭제</span></li>
			<% END IF %>


          <!--li id="deleteone"><span onClick="deleteWork(true)">영구삭제</span></li>
		  <img src="../../Home/images/i_bar.gif" align="absmiddle">
          <li id="deleteall" style="display:none"><span onClick="delAllFile()">모두삭제</span></li>
          <img src="../../Home/images/i_bar.gif" align="absmiddle">
          <li><span onClick="refresh_onclick()">새로고침</span></li>
          <li><span onClick="prevShow_onclick()">미리보기</span></li>
          
          <li><span onClick="Received_MailALLD()">모두삭제</span></li>
          
		  <li id="receivecheck" style="display:none" ><span onClick="receiveCheck_onClick()">수신확인</span></li>
          <li id="btnReject" style="display:none"><span onClick="reject_onclick()">수신거부</span></li-->
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;











          <li style="background:none;padding:0">          
            <select name="select" onChange="on_changeView(select.value)" style="WIDTH:110px">       
				
				<% if QSelect = "제목" then %>
				  <option VALUE="제목" selected>제목</option>
				<% else %>
				  <option VALUE="제목">제목</option>
				<% end if %>
				
				<% if QSelect = "보낸이" then %>
				  <option VALUE="보낸이" selected>보낸이</option>
				<% else %>
				  <option VALUE="보낸이">보낸이</option>
				<% end if %>
				
				<% if QSelect = "현장명" then %>
				  <option VALUE="현장명" selected>현장명</option>
				<% else %>
				  <option VALUE="현장명">현장명</option>
				<% end if %>

				<% if QSelect = "받는사람" then %>
				  <option VALUE="받는사람" selected>받는사람</option>
				<% else %>
				  <option VALUE="받는사람">받는사람</option>
				<% end if %>

				<% if QSelect = "공문번호(대,발)" then %>
				  <option VALUE="공문번호(대,발)" selected>공문번호(대,발)</option>
				<% else %>
				  <option VALUE="공문번호(대,발)"> 공문번호(대,발)</option>
				<% end if %>

				<% if QSelect = "공문번호(책,감)" then %>
				  <option VALUE="공문번호(책,감)" selected>공문번호(책,감)</option>
				<% else %>
				  <option VALUE="공문번호(책,감)">공문번호(책,감)</option>
				<% end if %>

				<% if QSelect = "공문번호(유,기)" then %>
				  <option VALUE="공문번호(유,기)" selected>공문번호(유,기)</option>
				<% else %>
				  <option VALUE="공문번호(유,기)">공문번호(유,기)</option>
				<% end if %>

              <!--option VALUE="읽지않은공문">읽지않은공문</option-->
            </select>
          </li>

          
		   <input type='textbox' size ="15" name='Gubun_i' VALUE="<%=Qgubun%>" onkeypress="javascript : if (event.keyCode == 13) goToPage1('<%=curpage%>');">
		  
           <li><span onClick="goToPage1(<%=curpage%>);" >검색</span></li>

        </ul>
      </div>
	  












<table class="mainlist" id ='test'>



  <form name="frmOutbox" action="BoardItemList.aspx" method="post">
    
	<tr>



<!--form action='#' method='post' name='form1'>
<input type='checkbox' name='cb1[]' value='cb1 1'>
<input type='checkbox' name='cb1[]' value='cb1 2'>
<input type='checkbox' name='cb1[]' value='cb1 3'>
<input type='checkbox' name='cb1[]' value='cb1 4'>
<input type='checkbox' name='cb1[]' value='cb1 5'>
<a id='CheckAll_cb1' href=''>Check All</a> | <a id='UncheckAll_cb1' href=''>Uncheck All</a></p-->



      <Th width=20 >
		<input type='checkbox' name="cbox" onclick='initCheckBehavior()' >
	  </Th>

      <!--th width=50 >번호</th-->

	  <% IF db_level = "S" THEN %>
		  <th width=20 ></th>
	  <% END IF %>
      
      <th style="cursor:hand;" width="300px" >제목</th>
      
      <th style="cursor:hand;" width="80px" >보낸이</th>
      
      <th style="cursor:hand;" width="100px" >현장명</th>
      
      <th style="cursor:hand;" width="80px" >보낸날</th>

      <!--th style="cursor:hand;" width="30px" >상태</th-->
      
      <th style="cursor:hand;padding:0" align="center" width="30px" ><img src="../../Home/images/file.gif" width="13" height="12"></th>
      


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
%>
		<% 'visited = DbRec("o_visited")
		  'if visited = 0 then
			'	sState = "안읽음"          
		  'else
		'		sState = "읽음"		  
		 ' end if%>

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


						if type1__ = "1" or type1__ = "2" THEN
							kname = DbRec("private_name")
							HJname =  DbRec("private_number")
						else 
							kname = DbRec("o_send_name")
							HJname = DbRec("o_send_longname")
						end if	
							
						%>



						<TR>
							
							<TD >
								<!--input type='checkbox' name='<%=qr%>' id='chk'-->
								<input type='checkbox' name='ccBox' id='chk' value="<%=DbRec("o_seq")%>" >
								<input type="hidden" name="db_acc" 		value="<%=DbRec("o_send_id")%>">
							</td>

							<!--TD style="cursor:hand;"><%=DbRec("o_seq")%></td-->
							<% IF db_level = "S" THEN %>
								<TD style="cursor:hand;">
									<% 'IF DbRec("S_II") = "" THEN %>
										<!--IMG name="Rec_Img" SRC="../ezPortal/Home/images/New.gif" border="0"-->
									<% 'else %>
										<IMG name="Rec_Img" SRC="../ezPortal/Home/images/New_empty.gif" border="0">
									<% 'end if %>
								</td>
							<% END IF %>

							<TD title='' style='cursor:hand;text-overflow:ellipsis; overflow:hidden;' 
							onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_send_id")%>","관리","<%=DbRec("o_send_longname")%>","<%=qq%>")'><nobr><%=DbRec("o_subject")%></nobr>
							</TD>

							<TD style="cursor:hand;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_send_id")%>","관리","<%=DbRec("o_send_longname")%>","<%=qq%>")'><%=kname%>
							</TD>

							<TD style="cursor:hand;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_send_id")%>","관리","<%=DbRec("o_send_longname")%>","<%=qq%>")'><%=HJname%>
							</TD>

							<TD style="cursor:hand;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_send_id")%>","관리","<%=DbRec("o_send_longname")%>","<%=qq%>")'><%=send_date%> 
							</TD>

							<!--TD name="State" id="State" style="cursor:hand;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>",)' onClick='ItemRead_onclick1("<%=DbRec("o_seq")%>","<%=HJName%>","<%=visited%>","<%=Aw%>")'>
							
							<div id='<%=Aw%>' ><%=sState%></div>
							
							</TD-->

							<TD style="cursor:hand;" onDblClick='ItemRead_onclick("<%=DbRec("o_seq")%>","<%=DbRec("o_send_id")%>","관리","<%=DbRec("o_send_longname")%>","<%=qq%>")'>

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


						</TR>

		
	<%					qq = qq + 1

		DbRec.MovePrevious
	Next
	%>


<% end if %>


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
<% set DbRec = Nothing %>
<% set DbRec2 = Nothing %>
<% set Result = Nothing %>
<% set Result1 = Nothing %>
<% set rs = Nothing %>
</body>
</HTML>



