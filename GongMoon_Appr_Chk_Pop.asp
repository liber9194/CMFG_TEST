<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<!--#include file="../../../../default_properties.asp" -->
<%
	db_id 	 	= session("db_id")
	db_level 	= session("db_level")
	db_level1 	= session("db_level1")
	site_code 	= session("site_code")

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >

<HTML>
<HEAD>
<title><%=HGubun%> </title>

 <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ks_c_5601-1987">

<meta content="Microsoft Visual Studio 7.0" name="GENERATOR">
<meta content="C#" name="CODE_LANGUAGE">
<meta content="JavaScript" name="vs_defaultClientScript">
<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
<LINK href="../../Home/css/default.css" type="text/css" rel="stylesheet">
<script language="JScript" src="/lang/ezEmail_ko.js"></script>
<script language="JavaScript" src="../../Home/myoffice/common/mouseeffect.js"></script>
<STYLE> 
P { MARGIN-BOTTOM: 0mm; MARGIN-TOP: 0mm } 
</STYLE>
<script language="JScript" src="js/reademail.js"></script>
<script language="JScript" src="js/string_component.js"></script>
<script>
	// 수정(2007.04.24) : exchange 버전별 처리
	var g_ExchangeVS = "2007";
	
	var g_paramURL = "http://exmail/exchange/204112/%EB%B0%9B%EC%9D%80%20%ED%8E%B8%EC%A7%80%ED%95%A8/[%ED%81%B4%EB%A6%B0%EC%8A%A4%ED%8C%B8]%20%EC%8A%A4%ED%8C%B8%EB%A9%94%EC%9D%BC%20%EB%82%B4%EC%97%AD%EC%9E%85%EB%8B%88%EB%8B%A4.(2009-05-31).EML";
	var g_expath = "exchange";
	var g_servername = "gw.dohwa.co.kr";
	var g_userID = "204112";
	var g_loginID = "204112";
	var g_author = "Basic RE9IV0EuQ08uS1JcMjA0MTEyOno4ODk1MjQ5";
	var g_exchNBName = "EXMAIL";
	var g_userName = "이재훈";
	var g_fromEmail = "root@cleanspam.dohwa.co.kr";
	var g_rejectWord = "210.122.146.203";
	var g_cancelsendread = "1";
	var g_notiSSO = "0";
	
	function window.onload()
	{
		window.onresize();
		
		if (g_notiSSO == "1")
			HideMenu();
	}

	function window.onbeforeprint() 
	{
		printScreen.style.display = "";
		normalScreen.style.display = "none";
		AttachFile.style.display = "none";
		parentBody.className = "";
		
		printMsgFrom.innerHTML = MsgToPut.innerHTML;
		printMsgTo.innerHTML = MsgToGot.innerHTML;
		printMsgCC.innerHTML = MsgCCGot.innerHTML;
		printSubject.innerHTML = mailSubject.innerHTML;
		printInsertFile.innerHTML = attachedfileDIV.innerHTML;
		printDocument.innerHTML = message.innerHTML;

		var checks = printInsertFile.all.tags("input");
		for (var i=0; i<checks.length; i++)
			checks.item(i).style.display = "none";

		var tableColl = printDocument.all.tags("TABLE");
		for (var i=0; i<tableColl.length; i++)
		{
			if (String(tableColl.item(i).borderColorDark).toLowerCase() == "#ffffff")
			{
				tableColl.item(i).style.borderCollapse = "collapse";
				tableColl.item(i).borderColorDark = "black";
			}
		}
	}

	function window.onafterprint() 
	{
		printScreen.style.display = "none";
		AttachFile.style.display = "";
		normalScreen.style.display = "";
		parentBody.className = "popup";
	}

	function window.onresize()
	{
		if (g_notiSSO == "1")
			return;
			
		if ( document.all.message.length > 1)
		{
			if (document.all.message(0).style.width != document.body.clientWidth - 20)
				document.all.message(0).style.width = document.body.clientWidth - 20;
			
		}
		else
		{
			if (document.all.message.style.width != document.body.clientWidth - 20)
				document.all.message.style.width = document.body.clientWidth - 20;
		}
	}	
	
	function HideMenu()
	{
		btnReply.style.display = "none";
		btnAllReply.style.display = "none";
		btnForward.style.display = "none";
		btnMove.style.display = "none";
		btnDelete.style.display = "none";
		btnEncode.style.display = "none";
		btnBoard.style.display = "none";				
		btnBookmark.style.display = "none";
		btnViewWeb.style.display = "none";
		btn_KMS.style.display = "none";
		btnInsertAddr.style.display = "none";
	}
	
	function ToKMS()
	{			
		var url = "http://" + document.location.hostname + "/myoffice/ezKMS/kasset/KAssetConvert.aspx?Mode=new&Flag=email&url="+ escape(g_paramURL);
	
		var feature = "status:no;dialogWidth:700px;dialogHeight:700px;help:no;scroll:no;edge:sunken";
		var RtnVal = window.showModalDialog(url, "", feature);						
	}
	
	function OnBtnClose()
	{

		//alert(document.all.end_date.value);

		if ((document.all.end_date.value == "") || (document.all.end_date.value == " ")){

			if(document.all.sort_.value == "") {
		          alert("해당 일시를 넣어주셔야 저장이 됩니다");
			} else {
			   window.location.href  = "josa_1.asp?end_date=" + document.all.end_date.value + "&sort_=" + document.all.sort_.value;				
			}
		} else {
			   window.location.href  = "josa_1.asp?end_date=" + document.all.end_date.value + "&sort_=" + document.all.sort_.value;				
		}

		//if (g_notiSSO == "1")
		//	window.location = "btn:action|close";
		//else
		//	window.close();
	}	
	
function printpr()
{
    var ezUtil = new ActiveXObject("ezUtil.MiscFunc");
			ezUtil.PrintPreview(document);
			ezUtil = null;
}






		function ItemRead_onclick(pItemBoardID,rid,aab)
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
			//alert(pItemBoardID);
			//alert(rID);
			//alert(Stype);
			//alert(HJname);
//			    window.open("mail_Result.asp?Seq=" + pItemBoardID + "&rID=" + rID + "&stype=" + Stype + "&HJname=" + HJname , "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,height=720,width=765,top=" + pTop + ",left=" + pLeft, "");	

			//alert(pItemBoardID);
			//alert(rid);


			window.location.href = "GongMoon_sujung.asp?Seq=" + pItemBoardID + "&code=" + rid + "&HGubun=" + aab + "&Ktype=<%=Ktype%>" ;

            //}
            //else
            //{
            //    window.open("BoardItemView.aspx?ShowAdjacent=" + ShowAdjacent + "&ItemID=" + pItemID + "&BoardID=" + pItemBoardID, "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=1,resizable=1,height=720,width=765,top=" + pTop + ",left=" + pLeft, "");
            //}
		}









		function Sujung_New(pItemBoardID,rid,aab)
		{

			window.location.href = "GongMoon_sujung_New_ag.asp?Seq=" + pItemBoardID + "&code=" + rid + "&HGubun=" + aab + "&Ktype=<%=Ktype%>" ;

		}




		function Test_Down11()
		{

			//alert("q");
			winstyle= "height=420,width=445, status=no,toolbar=no,menubar=no,location=no"
			window.open("../Down_ActiveX/DownLoad_ActiveX.asp?Seq=<%=Seq%>&code=<%=code%>",null,winstyle);

			//window.open("../Down_ActiveX/1.asp?Seq=<%=Seq%>&code=<%=code%>",null,winstyle);


			//window.location.href = "GongMoon_sujung.asp?Seq=" + pItemBoardID + "&code=" + rid + "&HGubun=" + aab + "&Ktype=<%=Ktype%>" ;

		}









		function Test_Down110()
		{

			//alert("q");
			winstyle= "height=420,width=445, status=no,toolbar=no,menubar=no,location=no"
			window.open("../Down_ActiveX/DownLoad_ActiveX_ag.asp?Seq=<%=Seq%>&code=<%=code%>",null,winstyle);

			//window.location.href = "GongMoon_sujung.asp?Seq=" + pItemBoardID + "&code=" + rid + "&HGubun=" + aab + "&Ktype=<%=Ktype%>" ;

		}









 function bdColor(){



  if(arguments[0] == 'blur'){
		//alert("1");

	   this.style.borderColor="#C0C0C0 #C0C0C0 #C0C0C0 #C0C0C0";
	   if (arguments[1] == '용역명'){document.frmOutbox.prj_detail_name.value = document.frmOutbox.prjname.value }

	   if (arguments[1] == '사업비'){

			var chk_1 = true;
			var chk_2 = true;

			var h1 = '';
			var h2 = '';

			//alert("1");

				h1 = checkIsDate_N1(document.frmOutbox.prj_saup_money.value)

				if (h1 == '') {
					document.frmOutbox.prj_saup_money.value = 0;
				} else {
					document.frmOutbox.prj_saup_money.value = h1;
				}

				h2 = checkIsDate_N1(document.frmOutbox.prj_saup_money1.value)
				if (h2 == '') {
					document.frmOutbox.prj_saup_money1.value = 0;
				} else {
					document.frmOutbox.prj_saup_money1.value = h2;
				}
				document.frmOutbox.prj_saup_money2.value = Isbun1(Number(Isbun(document.frmOutbox.prj_saup_money.value)) + Number(Isbun(document.frmOutbox.prj_saup_money1.value)));

			/**if(document.frmOutbox.prj_saup_money.value == '' || document.frmOutbox.prj_saup_money1.value == ''){
				if (document.frmOutbox.prj_saup_money.value == '') {
					document.frmOutbox.prj_saup_money.style.background ="#FF0000 #FF0000 #FF0000 #FF0000";
				} else {
					document.frmOutbox.prj_saup_money1.style.background ="#FF0000 #FF0000 #FF0000 #FF0000";
				}
				document.frmOutbox.prj_saup_money2.style.background ="#FF0000 #FF0000 #FF0000 #FF0000";

			} else {

				if(checkIsStr(document.frmOutbox.prj_saup_money.value)){  //날짜가 맞을경우는 true 를 반환
					//alert('날짜형식이 맞군요');
					document.frmOutbox.prj_saup_money.style.background ="#ffffff #ffffff #ffffff #ffffff";
				}else{
					document.frmOutbox.prj_saup_money.style.background ="#FF0000 #FF0000 #FF0000 #FF0000";
					chk_1 = false;
					//alert('틀린 금액형식입니다.');					
				}

				if(checkIsStr(document.frmOutbox.prj_saup_money1.value)){  //날짜가 맞을경우는 true 를 반환
					//alert('날짜형식이 맞군요');
					document.frmOutbox.prj_saup_money1.style.background ="#ffffff #ffffff #ffffff #ffffff";
				}else{
					document.frmOutbox.prj_saup_money1.style.background ="#FF0000 #FF0000 #FF0000 #FF0000";
					chk_2 = false;
					//alert('틀린 금액형식입니다.');					
				}
				

				if (chk_1 == true && chk_2 == true) {

					//alert(Isbun(document.frmOutbox.prj_saup_money.value));
					//alert(Isbun(document.frmOutbox.prj_saup_money1.value));

					document.frmOutbox.prj_saup_money2.value = Isbun1(Number(Isbun(document.frmOutbox.prj_saup_money.value)) + Number(Isbun(document.frmOutbox.prj_saup_money1.value)));
					document.frmOutbox.prj_saup_money2.style.background ="#ffffff #ffffff #ffffff #ffffff";
				} else {
					document.frmOutbox.prj_saup_money2.style.background ="#FF0000 #FF0000 #FF0000 #FF0000";
				}

			}
			**/
	   }	


	   if (arguments[1] == 'Last'){

			var h1 = '';

			//alert("1");

				h1 = checkIsDate_N1(document.frmOutbox.month_01.value)

				if (h1 == '') {
					document.frmOutbox.month_01.value = 0;
				} else {
					document.frmOutbox.month_01.value = h1;
				}

				h1 = checkIsDate_N1(document.frmOutbox.month_02.value)

				if (h1 == '') {
					document.frmOutbox.month_02.value = 0;
				} else {
					document.frmOutbox.month_02.value = h1;
				}

				h1 = checkIsDate_N1(document.frmOutbox.month_03.value)

				if (h1 == '') {
					document.frmOutbox.month_03.value = 0;
				} else {
					document.frmOutbox.month_03.value = h1;
				}

				h1 = checkIsDate_N1(document.frmOutbox.month_04.value)

				if (h1 == '') {
					document.frmOutbox.month_04.value = 0;
				} else {
					document.frmOutbox.month_04.value = h1;
				}

				h1 = checkIsDate_N1(document.frmOutbox.month_05.value)

				if (h1 == '') {
					document.frmOutbox.month_05.value = 0;
				} else {
					document.frmOutbox.month_05.value = h1;
				}

				h1 = checkIsDate_N1(document.frmOutbox.month_06.value)

				if (h1 == '') {
					document.frmOutbox.month_06.value = 0;
				} else {
					document.frmOutbox.month_06.value = h1;
				}

				h1 = checkIsDate_N1(document.frmOutbox.month_07.value)

				if (h1 == '') {
					document.frmOutbox.month_07.value = 0;
				} else {
					document.frmOutbox.month_07.value = h1;
				}

				h1 = checkIsDate_N1(document.frmOutbox.month_08.value)

				if (h1 == '') {
					document.frmOutbox.month_08.value = 0;
				} else {
					document.frmOutbox.month_08.value = h1;
				}

				h1 = checkIsDate_N1(document.frmOutbox.month_09.value)

				if (h1 == '') {
					document.frmOutbox.month_09.value = 0;
				} else {
					document.frmOutbox.month_09.value = h1;
				}

				h1 = checkIsDate_N1(document.frmOutbox.month_10.value)

				if (h1 == '') {
					document.frmOutbox.month_10.value = 0;
				} else {
					document.frmOutbox.month_10.value = h1;
				}

				h1 = checkIsDate_N1(document.frmOutbox.month_11.value)

				if (h1 == '') {
					document.frmOutbox.month_11.value = 0;
				} else {
					document.frmOutbox.month_11.value = h1;
				}

				h1 = checkIsDate_N1(document.frmOutbox.month_12.value)

				if (h1 == '') {
					document.frmOutbox.month_12.value = 0;
				} else {
					document.frmOutbox.month_12.value = h1;
				}

				
				
				document.frmOutbox.month_sum.value = Isbun1(Number(Isbun(document.frmOutbox.month_01.value))
												          + Number(Isbun(document.frmOutbox.month_02.value)) 
												          + Number(Isbun(document.frmOutbox.month_03.value)) 
												          + Number(Isbun(document.frmOutbox.month_04.value)) 
												          + Number(Isbun(document.frmOutbox.month_05.value)) 
												          + Number(Isbun(document.frmOutbox.month_06.value)) 
												          + Number(Isbun(document.frmOutbox.month_07.value)) 
												          + Number(Isbun(document.frmOutbox.month_08.value)) 
												          + Number(Isbun(document.frmOutbox.month_09.value)) 
												          + Number(Isbun(document.frmOutbox.month_10.value)) 
												          + Number(Isbun(document.frmOutbox.month_11.value)) 														  
												          + Number(Isbun(document.frmOutbox.month_12.value)))										  

	   }	







	   if (arguments[1] == '날짜체크'){
		   //document.frmOutbox.prj_detail_name.value = document.frmOutbox.prjname.value 
			if(this.value != ''){

				if(checkIsDate(this.value)){  //날짜가 맞을경우는 true 를 반환
					//alert('날짜형식이 맞군요');
					this.style.background ="#ffffff #ffffff #ffffff #ffffff";
					this.value = checkIsDate_N(this.value)
				}else{					
					
					//alert('틀린 날짜형식입니다.');
					if (this.value != '') {
						this.style.background ="#FF0000 #FF0000 #FF0000 #FF0000";
						this.value = checkIsDate_N(this.value)						

					}
					
					
				}
			}
	   }	


	   if (arguments[1] == '금액체크'){
			//alert(this.value);
			if(this.value == '' ){

			} else {

				if(checkIsStr(this.value)){  //날짜가 맞을경우는 true 를 반환
					//alert('날짜형식이 맞군요');
					this.style.background ="#ffffff #ffffff #ffffff #ffffff";
				}else{					
					//alert('금액형식이 틀리군요');
					//if (this.value != '') {
					//	this.style.background ="#FF0000 #FF0000 #FF0000 #FF0000";
					//}
					this.value = checkIsDate_N1(this.value)
							
				}

			}
	   }	


	   if (arguments[1] == '퍼센트체크'){
			//alert(this.value);
			if(this.value == '' ){

			} else {

				if(checkIsStr_dot(this.value)){  //날짜가 맞을경우는 true 를 반환
					//alert('날짜형식이 맞군요');
					this.value = Refire_1(this.value)
					//this.value = Per_Sent(this.value)
					
					this.style.background ="#ffffff #ffffff #ffffff #ffffff";
				}else{
					//alert('퍼센트체크 형식이 틀리군요');
					this.value = checkIsDate_N2(this.value)


					//this.value = Per_Sent(this.value)
					this.value = Refire_1(this.value)


					//if (this.value != '') {
					//	this.style.background ="#FF0000 #FF0000 #FF0000 #FF0000";
					//}
							
				}

			}
	   }	




  }else if(arguments[0] == 'focus'){
   this.style.borderColor="#FF0000 #FF0000 #FF0000 #FF0000";
  }
 }


function chk_1(aaa){
//alert(aaa.value.length);
//var val=document.getElementById(aaa);
//alert(aaa.value);
if(aaa.value.length==4||aaa.value.length==7){
aaa.value+='-';
}

}

function read_gongMoon(o_seq, o_receive_id, o_send_longname, o_visited, private__) {
	
	var pheight = window.screen.availHeight;
	var pwidth = window.screen.availWidth;
	var pTop = (pheight - 720) / 2;
	var pLeft = (pwidth - 765) / 2;
				
	// Right_Main_GongMoon.asp 에서 참고
	//window.open("../ezEmail/mail_read.asp?Seq=" + o_seq + "&rID=" + o_receive_id + "&stype=" + Stype + "&HJname=" + HJname + "&visited=" + visited + "&PP=" + PP , "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,height=700,width=775,top=" + pTop + ",left=" + pLeft, "");
	
	url = "../ezEmail/mail_read.asp?Seq=" + o_seq + "&rID=" + o_receive_id + "&stype=접수&HJname=" + o_send_longname + "&visited=" + o_visited + "&PP=" + private__;
	
	window.open(url,"", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,height=700,width=775,top=" + pTop + ",left=" + pLeft, "");	
}

</script>
</HEAD>

<body class="popup"> 
	<h1>
	<% if db_level = "S" or db_level = "Z" then %>
		PM 결재현황 [결재된 문서 : 후결재/결재올림 처리한 문서]
	<% elseif db_level = "P" then %>
		공문결재현황 [결재된 문서 : 후결재/결재올림 처리한 문서]
	<% end if %>
	</h1>
	<div id="close">
		<ul>
			<li onClick="window.close()"><span>닫기</span></li>
		</ul>
	</div>
	
	<% if db_level = "S" or db_level = "Z" then %>	
		<%
		Set rs=Server.CreateObject("ADODB.Recordset")
		rs.CursorType=1
	
		sql = " select a.c_userid, c_name = replace(a.c_name, ' ', ''), pm = replace(a.c_name, ' ', '') + ' (' + a.c_userid + ')' "
		
		if Request.ServerVariables("http_host") = g_domain then
			sql = sql & " from [" & g_cmfgDB & "].cug_Test.dbo.user_tbl a "
		else
			sql = sql & " from sfg.cug_Test.dbo.user_tbl a "
		end if
				
		sql = sql & " 	inner join dh_sap.dbo.tbl_sap_emp_info b on a.c_userid = b.emp_id "
		sql = sql & " where a.c_level = 'P' "
		sql = sql & " 	and b.RETIRE_DT > (select format(getdate(), 'yyyyMMdd')) "
		
		if search_txt <> "" then
			if search = "pm_id" then
				sql = sql & " and a.c_userid = '" & search_txt & "' "
			elseif search = "pm_name" then
				sql = sql & " and replace(a.c_name, ' ', '') = '" & search_txt & "' "
			end if
		end if
		
		rs.Open sql, DbCon_Mis
		%>
		
		<div class="box" style="height:600">
			<table width="100%" class="popuplist">
				<th style="text-align:center;">No</th>
				<th style="text-align:center;">해당PM</th>
				<th style="text-align:center;">접수문서(건)</th>
				<th style="text-align:center;">읽음 문서(건)</th>
				<th style="text-align:center;">안읽음 문서(건)</th>
				<th style="text-align:center;">결재된 문서(건)</th>
				
				<%
				if rs.Recordcount <> 0 then
					cnt = 0
					for i = 0 to rs.Recordcount - 1
					
						Set DbRec=Server.CreateObject("ADODB.Recordset")
						DbRec.CursorType=1
			
						sql = "			select count(*) as 'receive_cnt', "
						sql = sql & " 		count(case when b.o_visited > 0 then 1 end) as 'read_cnt', "
						sql = sql & " 		count(case when b.o_visited = 0 then 1 end) as 'unread_cnt', "
						sql = sql & " 		count(case when b.Result_Add = '1' or b.Result_Add = '3' then 1 end) as 'appr_cnt' " '후결재, 결재올림을 결재된 문서로 인정
						sql = sql & " 	from office_tbl a "
						sql = sql & " 		inner join receive_tbl b on a.o_seq = b.o_seq and b.receive_del = '' "
						sql = sql & " 	where b.o_receive_id = '" & rs("c_userid") & "' "
						sql = sql & " 		and a.office_del = '' "
						sql = sql & " 		and b.o_rdel_flag = 0 "
				
						DbRec.Open sql, DbCon
						
						if DbRec.Recordcount <> 0 then
							cnt = cnt + 1
				%>
							<tr>
								<td style="text-align:center;padding:5px;"><%=cnt%></td>
								<td style="text-align:center;padding:5px;"><%=rs("pm")%></td>
								<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("receive_cnt"),0,-1,0,-1)%></td>
								<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("read_cnt"),0,-1,0,-1)%></td>
								<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("unread_cnt"),0,-1,0,-1)%></td>
								<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("appr_cnt"),0,-1,0,-1)%></td>
							</tr>
				<%		 	end if 
							set DbRec = Nothing 
						rs.MoveNext
						next
					end if
					set rs = Nothing
				%>
				
			</table>
		</div>
		
	<% elseif db_level = "P" then %>
	
		<%			
		Set DbRec=Server.CreateObject("ADODB.Recordset")
		DbRec.CursorType=1
	
		sql = "			select count(*) as 'receive_cnt', "
		sql = sql & " 		count(case when b.o_visited > 0 then 1 end) as 'read_cnt', "
		sql = sql & " 		count(case when b.o_visited = 0 then 1 end) as 'unread_cnt', "
		sql = sql & " 		count(case when b.Result_Add = '1' or b.Result_Add = '3' then 1 end) as 'appr_cnt' "
		sql = sql & " 	from office_tbl a "
		sql = sql & " 		inner join receive_tbl b on a.o_seq = b.o_seq and b.receive_del = '' "
		sql = sql & " 	where b.o_receive_id = '" & db_id & "' "
		sql = sql & " 		and a.office_del = '' "
		sql = sql & " 		and b.o_rdel_flag = 0 "
		
		DbRec.Open sql, DbCon

		if DbRec.Recordcount <> 0 then
		%>
			<div class="box" style="height:600">
				<table width="100%" class="popuplist">					
					<th style="text-align:center;">접수문서(건)</th>
					<th style="text-align:center;">읽음 문서(건)</th>
					<th style="text-align:center;">안읽음 문서(건)</th>
					<th style="text-align:center;">결재된 문서(건)</th>
					
					<tr>
						<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("receive_cnt"),0,-1,0,-1)%></td>
						<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("read_cnt"),0,-1,0,-1)%></td>
						<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("unread_cnt"),0,-1,0,-1)%></td>
						<td style="text-align:right;padding:5px;"><%=formatnumber(DbRec("appr_cnt"),0,-1,0,-1)%></td>
					</tr>
				</table>
			</div>
			<% set DbRec = Nothing %>
		<% end if %>	
	<% end if %>
	
</body>
</HTML>
