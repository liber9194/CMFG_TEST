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

	if not db_id = "" then
		'if db_id = "216050" or db_id = "206171" then
			SQL = ""
			SQL = SQL & " SELECT * "
			SQL = SQL & " FROM "
			SQL = SQL & " ( "
			SQL = SQL & " 	( "
			SQL = SQL & " 		select aa.o_seq, aa.o_send_id, aa.o_send_name, aa.o_subject, aa.o_send_date, aa.o_send_longname, bb.o_receive_id, bb.o_receive_longname, bb.o_visited, aa.private__, aa.type__, aa.action_copy "
			
			if Request.ServerVariables("http_host") = g_domain then
				SQL = SQL & " 		from [" & g_cmfgDB & "].cug_test.dbo.office_tbl aa "
				SQL = SQL & " 			right join [" & g_cmfgDB & "].cug_test.dbo.receive_tbl bb on aa.o_seq = bb.o_seq "
			else
				SQL = SQL & " 		from sfg.cug_test.dbo.office_tbl aa "
				SQL = SQL & " 			right join sfg.cug_test.dbo.receive_tbl bb on aa.o_seq = bb.o_seq "
			end if			
			
			SQL = SQL & " 		where aa.private__ = '' and aa.o_send_id like '______' " ' len(o_send_id) = 6 => 공문발송함에서 보낸 공문만, 본사 관리자에서 현장에게 보낸 것만
			SQL = SQL & " 			and datediff(month, aa.o_send_date, getdate()) <= 6 "
			SQL = SQL & " 			and ((aa.type__ = '3' and aa.action_copy <> '') or aa.type__ <> '3') "
			SQL = SQL & " 	) a "
			SQL = SQL & " 	INNER JOIN "
			SQL = SQL & " 	( "
			SQL = SQL & " 		SELECT aa.project_code, aa.emp_num, aa.enter_date, aa.ret_date, bb.project_name "
			SQL = SQL & " 		FROM mis9803.dbo.[Tbl_Dispatch] aa "
			SQL = SQL & " 			Inner join mis9803.dbo.VISite_Project bb on aa.[project_code] = bb.[project_code] "
			'SQL = SQL & " 		WHERE aa.emp_num = '206171' "			
			SQL = SQL & " 		WHERE aa.emp_num = '" & db_id & "' "
			'SQL = SQL & " 	) b on a.o_receive_id = b.project_code and a.o_send_date between b.enter_date and b.ret_date "
			SQL = SQL & " 	) b on a.o_receive_id = b.project_code and ((b.ret_date is not null and (a.o_send_date between b.enter_date and b.ret_date)) or (b.ret_date is null and a.o_send_date >= b.enter_date)) "
			SQL = SQL & " 	left join "
			SQL = SQL & " 	( "
			SQL = SQL & " 		SELECT read_o_seq = o_seq, read_o_receive_id = o_receive_id, read_emp_num = emp_num "
			
			if Request.ServerVariables("http_host") = g_domain then
				SQL = SQL & " 		FROM [" & g_cmfgDB & "].cug_test.dbo.GongMoon_Chk "
			else
				SQL = SQL & " 		FROM sfg.cug_test.dbo.GongMoon_Chk "
			end if
						
			SQL = SQL & " 	) chk on a.o_seq = chk.read_o_seq and a.o_receive_id = chk.read_o_receive_id and b.emp_num = chk.read_emp_num "
			SQL = SQL & " ) "
			SQL = SQL & " WHERE read_o_seq is null "
			SQL = SQL & " ORDER BY o_send_date desc "
			
			Set rs_gongMoon_chk = Server.CreateObject("ADODB.Recordset")
			rs_gongMoon_chk.CursorType=1
			
			rs_gongMoon_chk.Open SQL, DbCon_Mis			
			
		'end if
	end if


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

function read_gongMoon(o_seq, o_receive_id, o_send_longname, o_visited, private__, type__) {
	
	var pheight = window.screen.availHeight;
	var pwidth = window.screen.availWidth;
	var pTop = (pheight - 720) / 2;
	var pLeft = (pwidth - 765) / 2;
				
	// Right_Main_GongMoon.asp 에서 참고
	//window.open("../ezEmail/mail_read.asp?Seq=" + o_seq + "&rID=" + o_receive_id + "&stype=" + Stype + "&HJname=" + HJname + "&visited=" + visited + "&PP=" + PP , "", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,height=700,width=775,top=" + pTop + ",left=" + pLeft, "");
	
	if (type__ == "3")
	{
		url = "../ezEmail/mail_read_CM.asp?Seq=" + o_seq + "&rID=" + o_receive_id + "&stype=접수&HJname=" + o_send_longname + "&visited=" + o_visited + "&PP=" + private__;
	}
	else
	{
		url = "../ezEmail/mail_read.asp?Seq=" + o_seq + "&rID=" + o_receive_id + "&stype=접수&HJname=" + o_send_longname + "&visited=" + o_visited + "&PP=" + private__;
	}	
	
	window.open(url,"", "toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,height=700,width=775,top=" + pTop + ",left=" + pLeft, "");	
}

</script>
</HEAD>

<body class="popup"> 
	<h1> 공문 리스트 </h1>
	<div id="close">
		<ul>
			<li onClick="window.close()"><span>닫기</span></li>
		</ul>
	</div>
	
	<% if rs_gongMoon_chk.Recordcount = 0 then %>
		<h2>미조회 항목이 없습니다.</h2>
	<% else %>
	
		<div class="box" style="height:600">
			<table width="100%" class="popuplist">
				<th>제목</th>
				<th>보낸이</th>
				<th>발신현장명</th>
				<th>보낸날</th>
				<th>조회여부</th>
				
				<%
					if rs_gongMoon_chk.Recordcount > 0 then
						for i = 1 to rs_gongMoon_chk.Recordcount
				%>
							<tr>
								<td><a href="javascript:read_gongMoon('<%=rs_gongMoon_chk("o_seq")%>', '<%=rs_gongMoon_chk("o_receive_id")%>', '<%=rs_gongMoon_chk("o_send_longname")%>', '<%=rs_gongMoon_chk("o_visited")%>', '<%=rs_gongMoon_chk("private__")%>', '<%=rs_gongMoon_chk("type__")%>');"><%=rs_gongMoon_chk("o_subject")%></td>
								<td style="text-align:center;"><%=rs_gongMoon_chk("o_send_name")%></td>
								<td style="text-align:center;"><%=rs_gongMoon_chk("o_send_longname")%></td>
								<td style="text-align:center;"><%=rs_gongMoon_chk("o_send_date")%></td>
								<td style="text-align:center;">
								<%
									if not isnull(rs_gongMoon_chk("read_o_seq")) then
										response.write "읽음"
									else
										response.write "안읽음"
									end if
								%>
								</td>
							</tr>
				<%			
							rs_gongMoon_chk.MoveNext
						next
					end if
				%>
				
			</table>
		</div>
	
	<% end if %>
	
</body>
</HTML>
