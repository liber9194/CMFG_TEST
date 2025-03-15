<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="misop.inc" -->
<%
db_id 	 	= session("db_id")
db_level 	 	= session("db_level")
db_level1 	 	= session("db_level1")
site_code 	 	= session("site_code")
site_name 	 	= session("site_name")

Qang   = request("Qang")

if Qang = "" then
	Qang = 0
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
<HEAD>
	<title>left_myoffice</title>
	<meta name="GENERATOR" Content="Microsoft Visual Studio 7.0">
	<meta name="CODE_LANGUAGE" Content="C#">
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<meta name="vs_defaultClientScript" content="JavaScript">
	<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	<link rel="stylesheet" href="../../Home/css/default.css" type="text/css">
	<link rel="stylesheet" href="../../Home/css/email_tree.css" type="text/css">
	<link href="../../Home/skin/skin_1/skin.css" rel="stylesheet" type="text/css">


	<script language="JScript"  src="../ezEmail/lang/ezEmail_ko.js" ></script>
	<script language="JScript" src="../ezEmail/js/email_tree.js"></script>
	
	<script language="JavaScript" src="../../Home/myoffice/common/mouseeffect.js"></script>

<script>
	// 수정(2007.02.07) : WebDAV 루틴 변경 (BE 서버)
	// 수정(2007.04.24) : exchange 버전별 처리
	var g_ExchangeVS = "2007";
	var g_szMailFolderURL;
	if( g_ExchangeVS == "2007" )
		g_szMailFolderURL = "http://EXMAIL/exchange/204112";
	else
		g_szMailFolderURL = "http://gw.dohwa.co.kr/exchange/204112";
	
	var g_firstOpen = true;



	// 지정된 기능을 실행시킨다.
	function window.onload()
	{
		Function_Flag('1', 1);
	}
	// 메일쓰기
	function write_Letter(aa1)
	{
		parent.frames[2].location = "right_main_gongmoon.asp?type__=" + aa1 ;		
	}
	
	function write_Letter_CM(aa1)
	{
		parent.frames[2].location = "right_main_gongmoon_CM.asp?type__=" + aa1 ;		
	}

	function write_Letter_CM_test(aa1)
	{
		parent.frames[2].location = "right_main_gongmoon_CM_testDB.asp?type__=" + aa1 ;
	}
	
	function Ilban_read(aa1)
	{
		parent.frames[2].location = "ilban_gongmoon_read.asp?type__=" + aa1 ;		
	}

	function Ilban_write(aa1)
	{
		parent.frames[2].location = "ilban_gongmoon_write.asp?type__=" + aa1 ;		
	}

	function Ilban_manager(aa1)
	{
		parent.frames[2].location = "ilban_gongmoon_manager.asp?type__=" + aa1 ;		
	}

	function folder_manage(aa1)
	{
		parent.frames[2].location = "right_main_gongmoon_insert.asp?type__=" + aa1;
	}
	
	function folder_manage_cm(aa1)
	{
		parent.frames[2].location = "right_main_gongmoon_insert_CM.asp?type__=" + aa1;
	}
	
	function appr_sheet(aa1)
	{
		parent.frames[2].location = "gongmoon_appr_sheet.asp?cond=" + aa1;
	}

	// 메일함 트리뷰와 관련된 함수들
	function LoadEmailTree()
	{
		PostTreeView.config = treeconfig;
		PostTreeView.source = "<tree><nodes>" + get_childXML(g_szMailFolderURL, true, true) + "</nodes></tree>";
		try{
			PostTreeView.update();
		} catch(e) {};
	}

	function requestdata()
	{
		var nodeIdx = window.event.nodeIdx;
		var childxml = get_childXML(PostTreeView.getvalue(nodeIdx, "href"), false, true)
		PostTreeView.putchildxml(nodeIdx, childxml);
	}

	function selectnode()
	{
		var nodeIdx = PostTreeView.selectedIndex;
		var href = PostTreeView.getvalue(nodeIdx, "href");
		var url = "/myoffice/ezEmail/mail_list.aspx?dispname=" + escape(PostTreeView.getvalue(nodeIdx, "foldername")) + "&url=" + escape(PostTreeView.getvalue(nodeIdx, "href"))
		
		if (g_firstOpen)
			g_firstOpen = false;
		else
			window.open(url, "right");

		get_unreadcount();

		try {
			window.top.frames("top").document.Script.change_menu(2, "<a href='/myoffice/main/index_myoffice.asp?funcode=1' target='main' class='n'>메일</a> > " +
			"<a href='/myoffice/main/index_myoffice.asp?funcode=1' target='main' class='n'>메일</a> > <a href='" + url + "' target='right' class='n'>" + PostTreeView.getvalue(nodeIdx, "foldername") + "</a>");
		} catch(e) {}	
	}

	function email_dragdrop()
	{
		var szCommand = (window.event.bctrl) ? "copy" : "move";
		var szSubCommand = window.event.command;

		if (szCommand == "move" && szSubCommand == "ViewMailListMove")
		{
			try {
				window.parent.frames("right").document.Script.move_on_dragdrop(PostTreeView.getvalue(event.nodeIdx, "href"));
			} catch(e) {}
		}
		else if (szCommand == "copy" && szSubCommand == "ViewMailListMove")
		{
			try {
				window.parent.frames("right").document.Script.copy_on_dragdrop(PostTreeView.getvalue(event.nodeIdx, "href"));
			} catch(e) {}
		}
	}

	var g_xmlUnread = null;
	function get_unreadcount()
	{
		if (g_xmlUnread != null)
			return;

		var strXml;
		strXml = "<?xml version='1.0' encoding='ks_c_5601-1987'?>" +
       				"<a:propfind xmlns:a='DAV:' xmlns:b='urn:schemas:httpmail:'>" +
               		"<a:prop>" +
               		"<b:unreadcount/>" +
               		"</></>"

		g_xmlUnread = new ActiveXObject("Microsoft.XMLHttp");
		try
		{
			// 수정(2007.04.24) : exchange 버전별 처리
			if( g_ExchangeVS == "2007" )
			{
				// 수정(2007.02.12) : WebDAV 루틴 변경 (BE 서버)
				var xmlDOM = new ActiveXObject("Microsoft.XMLDOM");

				var objRoot = xmlDOM.createNode(1,"DATA","");
				xmlDOM.appendChild(objRoot);

				var objNode = xmlDOM.createNode(1, "QUERY", "");
				var objCDate = xmlDOM.createCDATASection(strXml);
				objNode.appendChild(objCDate);
				objRoot.appendChild(objNode);
	            
				var objNode = xmlDOM.createNode(1, "CMD", "");
				objNode.text = "PROPFIND";
				objRoot.appendChild(objNode);
	            
				var objNode = xmlDOM.createNode(1, "URL", "");
				objNode.text = PostTreeView.getvalue(PostTreeView.selectedIndex, "href");
				objRoot.appendChild(objNode);
		        
				var objNode = xmlDOM.createNode(1, "DEPTH", "");
				objNode.text = "0";
				objRoot.appendChild(objNode);
		        
				g_xmlUnread.open("POST", "/myoffice/ezEmail/remote/mail_interwebdav.aspx", true);
				g_xmlUnread.onreadystatechange = get_unreadend;
				get_unreadend.href = PostTreeView.getvalue(PostTreeView.selectedIndex, "href");
				g_xmlUnread.send(xmlDOM.xml);
			}
			else
			{
				g_xmlUnread.Open("PROPFIND", PostTreeView.getvalue(PostTreeView.selectedIndex, "href"), true);
				g_xmlUnread.setRequestHeader("Content-Type", "text/xml");
				g_xmlUnread.setRequestHeader("Depth:", "0");	
				g_xmlUnread.onreadystatechange = get_unreadend;
				get_unreadend.href = PostTreeView.getvalue(PostTreeView.selectedIndex, "href");
				g_xmlUnread.Send(strXml);
			}
		}
		catch(e)
		{
			g_xmlUnread = null;
		}
	}

	function get_unreadend()
	{
		if (g_xmlUnread == null || g_xmlUnread.readyState != 4)
			return;

		if (g_xmlUnread.status >=200 && g_xmlUnread.status < 300)
		{
			var xmlDom = new ActiveXObject("Microsoft.XMLDom");
			xmlDom = g_xmlUnread.responseXML;
			var unreadcount = xmlDom.getElementsByTagName("d:unreadcount").item(0).text;
			var caption = PostTreeView.getvalue(PostTreeView.selectedIndex, "foldername");

			if (get_unreadend.href == PostTreeView.getvalue(PostTreeView.selectedIndex, "href"))
			{
				if (unreadcount == "0")
				{
					PostTreeView.putcaption(PostTreeView.selectedIndex, caption);
					PostTreeView.putstyle(PostTreeView.selectedIndex, "font-weight:normal;");
				}
				else
				{
					PostTreeView.putcaption(PostTreeView.selectedIndex, caption + "(" + unreadcount + ")");
					PostTreeView.putstyle(PostTreeView.selectedIndex, "font-weight:bold;");
				}
				xmlDom = null;
			}
		}

		g_xmlUnread = null;
	}

	// 외부 메일 확인
	function check_pop3()
	{
		window.showModalDialog("/myoffice/ezEmail/mail_getpop3.aspx", "check pop3", "dialogWidth:460px; dialogHeight:360px; scroll:no; status:no; help:no; scroll:no; edge:sunken");
	}

	// 메일 내보내기
	function mail_export()
	{
		//parent.frames[2].location = "right_main_all.asp";
		parent.frames[2].location = "right_main_gongmoon_all.asp";

	}




	function mail_export1()
	{
		//parent.frames[2].location = "right_main_all.asp";
		parent.frames[2].location = "right_main_gongmoon_all_new1.asp?type__=";

	}

	function mail_export2()
	{
		//parent.frames[2].location = "right_main_all.asp";
		parent.frames[2].location = "right_main_gongmoon_all_new1.asp?type__=1";

	}

	function mail_export3()
	{
		//parent.frames[2].location = "right_main_all.asp";
		parent.frames[2].location = "right_main_gongmoon_all_new1.asp?type__=2";

	}

	function mail_export4()
	{
		//parent.frames[2].location = "right_main_all.asp";
		parent.frames[2].location = "right_main_gongmoon_all_new1.asp?type__=&pOO=1";

	}
	
	function mail_export5()
	{
		//parent.frames[2].location = "right_main_all.asp";
		parent.frames[2].location = "right_main_gongmoon_all_new1.asp?type__=3";

	}


	// 메일함 PC 저장하기
	function mail_exportall()
	{
		var param = {"href":new Array(), "parent":new Object(), "url":new String()};
		param["name"] = PostTreeView.getvalue(PostTreeView.selectedIndex, "foldername");
		param["url"] = PostTreeView.getvalue(PostTreeView.selectedIndex, "href");
		param["parent"] = window.parent.frames("right");
		
		// 수정(2007.04.24) : exchange 버전별 처리
		param["exchangevs"] = g_ExchangeVS;
		
		window.showModalDialog("/myoffice/ezEmail/htm/mail_exportall.aspx", param, "dialogWidth:480px; dialogHeight:265px; scroll:no; status:no; help:no; scroll:no; edge:sunken");		
	}

	//메일 가져오기
	function mail_import()
	{
		var param = new Array();
		param["foldername"] = PostTreeView.getvalue(PostTreeView.selectedIndex, "foldername");
		param["folderpath"] = PostTreeView.getvalue(PostTreeView.selectedIndex, "href");
		param["parent"] = window;
		
		// 수정(2007.04.24) : exchange 버전별 처리
		param["exchangevs"] = g_ExchangeVS;
		
		window.showModalDialog("/myoffice/ezEmail/htm/mail_import.aspx", param, "dialogWidth:429px; dialogHeight:265px; scroll:no; status:no; help:no; scroll:no; edge:sunken");
	}

	// 메일함 관리

	
	// 각 기능을 불러오는 함수
	function Function_Flag(v_data, subfolder)
	{   
		v_data=parseInt(v_data);
		
		switch(v_data)
		{
			case 1:		// 메일
				LoadEmailTree();
				//TreeView_toggle(POST_DIV, Open_Mail, subfolder);
				if (typeof(subfolder) != "undefined")
					Open_Mail(subfolder);	
				else
					Open_Mail();			
				break;
		}
	}
	
	//토글함수
	function TreeView_toggle(TreeView, TreeFunc, subfolder)
	{
		if (TreeView.style.display == "none")
		{
			//POST_DIV.style.display = "none";
			
			//TreeView.style.display = "block";

			if (typeof(subfolder) != "undefined")
				TreeFunc(subfolder);	
			else
				TreeFunc();			
		}
		else
			TreeView.style.display = "none";
	}

	// 메일 기능 실행
	function Open_Mail(treeid)
	{
		PostTreeView.select(treeid);
	}

	function Open_Search()
	{
		try {			
			var url = "/myoffice/ezEmail/mail_search.aspx";
			window.open(url, "right");
		} catch(e) {}	
	}

	function Open_ReservationManage()
	{
		window.showModalDialog("/myoffice/ezEmail/mail_reservation.aspx", "", "dialogHeight:350px; dialogWidth:501px; status:no;scroll:auto; help:no; edge:sunken");
	}

	function Open_Restore()
	{
		var pheight = window.screen.availHeight;
		var pwidth = window.screen.availWidth;
		var pTop = (pheight - 500) / 2;
		var pLeft = (pwidth - 700) / 2;
	
		var name = PostTreeView.getvalue(PostTreeView.selectedIndex, "foldername");
		var path = PostTreeView.getvalue(PostTreeView.selectedIndex, "href");

		// 수정(2005.06.22) : 영구삭제메일복원 화면 팝업 방식 변경%>
		//window.showModalDialog("/myoffice/ezEmail/mail_restore_deleted.aspx?name=" + escape(name) + "&path=" + escape(path), "", "dialogHeight:440px; dialogWidth:535px; status:no;scroll:auto; help:no; edge:sunken");
		//LoadEmailTree();
		window.open("/myoffice/ezEmail/mail_restore_deleted.aspx?name=" + escape(name) + "&path=" + escape(path), "", "width=700, height=425, status = no, toolbar=no, menubar=no, location=no, resizable=1, top=" + pTop + ",left=" + pLeft, "");
	}

	function Change_MailAddress()
	{
		window.open("/myoffice/ezOrgan/admin/ConfigEmail.aspx?id=" + "204112", "", "height=305px,width=430px,status=no,toolbar=no,menubar=no,location=no,resizable=1");
		//window.open("/myoffice/ezOrgan/admin/configquota.aspx?id=" + "204112", "", "height=290px,width=320px,status=no,toolbar=no,menubar=no,location=no,resizable=1");
	}
	

function win2(aaa){
//일반
//	w = window.open("select_id.asp?rad=All", 'id_list', "scrollbars=auto,width=700,height=450,left=300,top=150");

//관리자
//	alert("t");
	
	
	
	//
	//w = window.open("select_manager.asp", 'id_list', "scrollbars=auto,width=700,height=700,left=300,top=150");


	if("<%=db_id%>" == "204112") {
		w = window.open("../Addr_New/select_manager.asp", 'id_list', "scrollbars=auto,width=1000,height=700,left=0,top=0");
	} else {
		w = window.open("../Addr_New/select_manager.asp", 'id_list', "scrollbars=auto,width=1000,height=700,left=0,top=0");
		//w = window.open("../Addr/select_manager.asp", 'id_list', "scrollbars=auto,width=1000,height=700,left=0,top=0");
	}

	w.focus();

	return;

}


</script>
</HEAD>

<body class="leftbody" leftmargin="0" topmargin="0" rightmargin="0" style="OVERFLOW-Y:auto; OVERFLOW-X:auto">

	<xml id="treeconfig">
		<tree>
			<config>
				<size width="14" height="17" />
				<baseimage>
					<dot_continue path="/images/Email/tree/dot_continue.gif" />
					<dot_end path="/images/Email/tree/dot_end.gif" />
					<dot_normal path="/images/Email/tree/dot_normal.gif" />
					<minus_end path="/images/Email/tree/minus_end.gif" />
					<minus_normal path="/images/Email/tree/minus_normal.gif" />
					<plus_end path="/images/Email/tree/plus_end.gif" />
					<plus_normal path="/images/Email/tree/plus_normal.gif" />
					<space path="/images/Email/tree/space.gif" />
					<selected path="/images/Email/tree/folderselect.gif" />
				</baseimage>
				<baseclass>
					<normal name="node_normal" />
					<selected name="node_selected" />
					<hover name="node_hover" />
				</baseclass>
				<images>
					<image idx="1" path="/images/Email/tree/folder.gif" />
				</images>
			</config>
		</tree>
	</xml>
	<div id="left">
				
		
<% if UCASE(db_level) = "S" THEN %>
		<div class="left_mail" title="마이오피스"></div>
		<!--iframe width=100% height="110px" border=0 src='/myoffice/ezPortal/filter/URLPortlet.aspx?uid=c1488e35-c011-4906-b8d2-b6cd28cbc94e' frameborder=0 scrolling=no></iframe-->
		<h2>공문관리함</h2>
		<ul id="tree">
			<!--li evt="0"><span onClick="mail_export()" style="width:100%">공문관리함</span></li-->		
			<li evt="0"><span onClick="mail_export1()" style="width:100%">공문관리함<br>(대표이사->발주청)</span></li>				
			<li evt="0"><span onClick="mail_export2()" style="width:100%">공문관리함<br>(건설사업관리단->CM부문)</span></li>				
			<li evt="0"><span onClick="mail_export3()" style="width:100%">공문관리함<br>(유관기관,기타)</span></li>				
			<li evt="0"><span onClick="mail_export4()" style="width:100%">공문관리함<br>(CM부문->건설사업관리단)</span></li>
			<li evt="0"><span onClick="mail_export5()" style="width:100%">공문관리함<br>(CM부문(대표이사)->건설사업관리단, 유관기관 등)</span></li>
		</ul>
<% ELSEif UCASE(db_level) = "G" or UCASE(db_level) = "C" or UCASE(db_level) = "D" THEN  %>
		<div class="left_mail" title="마이오피스"></div>
		<!--iframe width=100% height="110px" border=0 src='/myoffice/ezPortal/filter/URLPortlet.aspx?uid=c1488e35-c011-4906-b8d2-b6cd28cbc94e' frameborder=0 scrolling=no></iframe-->
		<h2>공문접수함</h2>
		<ul>
		<% if UCASE(db_level) = "Z" or UCASE(db_level) = "P" THEN %>
					<li evt="0"><span onClick="write_Letter('')" style="width:100%">공문접수함<br>(대표이사->발주청)</span></li>
					<li><span onClick="write_Letter('1')" style="width:100%">공문접수함<br>(건설사업관리단->CM부문)</span></li>
					<li><span onClick="write_Letter('2')" style="width:100%">공문접수함<br>(유관기관,기타)</span></li>
		<% 'elseif UCASE(db_level) = "C" THEN %>			
			<% if site_code <> "25-0000-000" then %>			
			<% else %>			
			<% END IF %>
		<% else %>
				<% if site_code <> "" then %>
					<li evt="0"><span onClick="write_Letter('')" style="width:100%">공문접수함</span></li>	
				<% END IF %>
		<% END IF %>
		</ul>
		<% if UCASE(db_level1) = "A" THEN %>
			<h2>공문관리함</h2>
			<ul id="tree">
				<!--li evt="0"><span onClick="mail_export()" style="width:100%">공문관리함</span></li-->				

				<li evt="0"><span onClick="mail_export1()" style="width:100%">공문관리함<br>(대표이사->발주청)</span></li>				
				<li evt="0"><span onClick="mail_export2()" style="width:100%">공문관리함<br>(건설사업관리단->CM부문)</span></li>				
				<li evt="0"><span onClick="mail_export3()" style="width:100%">공문관리함<br>(유관기관,기타)</span></li>				
				<li evt="0"><span onClick="mail_export4()" style="width:100%">공문관리함<br>(CM부문->건설사업관리단)</span></li>
			</ul>	
		<% END IF %>
<% ELSE %>
		<div class="left_mail" title="마이오피스"></div>
		<!--iframe width=100% height="110px" border=0 src='/myoffice/ezPortal/filter/URLPortlet.aspx?uid=c1488e35-c011-4906-b8d2-b6cd28cbc94e' frameborder=0 scrolling=no></iframe-->
		<% if db_id = "208097" THEN 'db_id = "206056" THEN %>
			<h2>공문함</h2>
		<% ELSE %>
			<h2>공문접수함</h2>
		<% END IF %>
		<ul>
		<% if UCASE(db_level) = "Z" or UCASE(db_level) = "P" THEN %>
			<% if db_id = "208097" THEN 'db_id = "208097" THEN %>				
					<li><span onClick="write_Letter('1')" style="width:100%">공문접수함<br>(건설사업관리단->CM부문)</span></li>
					<li><span onClick="write_Letter('2')" style="width:100%">공문접수함<br>(대외,FAX,기타)</span></li>
					<li><span onClick="write_Letter_CM('3')" style="width:100%">공문접수함(CM부문(대표이사)->건설사업관리단, 유관기관 등)</span></li>
			<% else %>
					<li evt="0"><span onClick="write_Letter('')" style="width:100%">공문접수함<br>(대표이사->발주청)</span></li>
					<li><span onClick="write_Letter('1')" style="width:100%">공문접수함<br>(건설사업관리단->CM부문)</span></li>
					<li><span onClick="write_Letter('2')" style="width:100%">공문접수함<br>(유관기관,기타)</span></li>
					<!--li><span onClick="write_Letter('3')" style="width:100%">공문접수함<br>(유관기관,기타)</span></li-->
					
					<li><span onClick="write_Letter_CM('3')" style="width:100%">공문접수함(CM부문(대표이사)->건설사업관리단, 유관기관 등)</span></li>
					<% if db_id = "216050" then %>
						<li><span onClick="write_Letter_CM_test('3')" style="width:100%">공문접수함(CM부문(대표이사)->건설사업관리단, 유관기관 등)</span></li>
					<% end if %>
			<% end if %>
		<% else %>

				<% if site_code <> "" then %>
					<li evt="0"><span onClick="write_Letter('')" style="width:100%">공문접수함</span></li>
					<li evt="0"><span onClick="write_Letter_CM('3')" style="width:100%">공문접수함(CM부문(대표이사)->건설사업관리단, 유관기관 등)</span></li>
					
					
				<% END IF %>
		<% END IF %>

				</ul>  
				<h2>공문발송함</h2>
				<ul id="tree">
		<% if UCASE(db_level) = "Z" or UCASE(db_level) = "P" THEN %>
			<% if db_id = "208097" THEN 'db_id = "208097" THEN %>	
            	    <li evt="0"><span onClick="write_Letter('')" style="width:100%">공문발송함<br>(대표이사->발주청)</span></li>
					<li evt="0"><span onClick="folder_manage('')" style="width:100%">공문발송함</span></li>
					<li evt="0"><span onClick="folder_manage_cm('3')" style="width:100%">공문발송함(CM부문(대표이사)->건설사업관리단, 유관기관 등)</span></li>
			<% ELSE %>
					<li evt="0"><span onClick="folder_manage('')" style="width:100%">공문발송함</span></li>
					
					<!-- 공문발송함(CM본부(대표이사) -> 건설사업관리단, 유관기관 등) 추가 요청(2017.11.10 경익수 상무 요청) -->
					<li evt="0"><span onClick="folder_manage_cm('3')" style="width:100%">공문발송함(CM부문(대표이사)->건설사업관리단, 유관기관 등)</span></li>
					
			<% END IF %>
		<% else %>
				<% if site_code <> "" then %>
					<li evt="0"><span onClick="folder_manage('')" style="width:100%">공문발송함<br>(대표이사->발주청)</span></li>
							<li><span onClick="folder_manage('1')" style="width:100%">공문발송함<br>(건설사업관리단->CM부문)</span></li>
				<% END IF %>
		<% END IF %>

				</ul>  

				<h2>주소록</h2>
				<ul id="tree">
					<% if UCASE(db_level) = "Z" or UCASE(db_level) = "P"  THEN %>
						<li evt="0"><span onClick="win2('<%=db_level%>')" style="width:100%">주소록</span></li>	
					<% else %>
						<% if site_code <> "" then %>
							<!--li evt="0"><span onClick="win2('<%=db_level%>')" style="width:100%">주소록</span></li-->	
						<% END IF %>

					<% END IF %>
				</ul>  

		<% if UCASE(db_level) = "Z" THEN %>
				<h2>공문관리함</h2>
				<ul id="tree">
					<!--li evt="0"><span onClick="mail_export()" style="width:100%">공문관리함</span></li-->		
					<li evt="0"><span onClick="mail_export1()" style="width:100%">공문관리함<br>(대표이사->발주청)</span></li>				
					<li evt="0"><span onClick="mail_export2()" style="width:100%">공문관리함<br>(건설사업관리단->CM부문)</span></li>				
					<li evt="0"><span onClick="mail_export3()" style="width:100%">공문관리함<br>(유관기관,기타)</span></li>				
					<li evt="0"><span onClick="mail_export4()" style="width:100%">공문관리함<br>(CM부문->건설사업관리단)</span></li>	
					<li evt="0"><span onClick="mail_export5()" style="width:100%">공문관리함<br>(CM부문(대표이사)->건설사업관리단, 유관기관 등)</span></li>					
				</ul>  
		<% END IF %>
		


		<% 'if db_id = "204112" or db_id = "206171" then %>			
				<!--h2>공문 관리함</h2>
				<ul id="tree">
					<li evt="0"><span onClick="mail_export1()" style="width:100%">공문관리함<br>(대표이사->발주청)</span></li>				
					<li evt="0"><span onClick="mail_export2()" style="width:100%">공문관리함<br>(건설사업관리단->CM부문)</span></li>				
					<li evt="0"><span onClick="mail_export3()" style="width:100%">공문관리함<br>(유관기관,기타)</span></li>				
					<li evt="0"><span onClick="mail_export4()" style="width:100%">공문관리함<br>(CM부문->건설사업관리단)</span></li>
				</ul-->  
		<% 'end if %>


<% END IF %>






















<% 'if db_id = "204112" then %>
<% 'if db_id = "204112" or db_id = "195171" or db_id = "207241"  then %>

	<% if UCASE(db_level) = "S" THEN %>

				<h2>요청자료</h2>
				<ul id="tree">
					<li evt="0"><span onClick="Ilban_write('')" style="width:100%">요청자료발송함</span></li>	
					<li><span onClick="Ilban_manager()" style="width:100%">요청자료관리함</span></li>				
				</ul>  

	<% ELSEif UCASE(db_level) = "G" or UCASE(db_level) = "C" or UCASE(db_level) = "D" THEN  %>

			
				<h2>요청자료</h2>
				<ul>
					<li evt="0"><span onClick="Ilban_read('')" style="width:100%">요청자료접수함</span></li>	

				<% if UCASE(db_level1) = "A" THEN %>
						<li><span onClick="Ilban_manager()" style="width:100%">요청자료관리함</span></li>				
				<% END IF %>

				</ul>

	<% ELSE %>

			<% if UCASE(db_level) = "Z" or UCASE(db_level) = "P" THEN %>

				<h2>요청자료</h2>
				<ul>
					<li evt="0"><span onClick="Ilban_write('')" style="width:100%">요청자료발송함</span></li>	
					<% if UCASE(db_level) = "Z" THEN %>				
								<li><span onClick="Ilban_manager()" style="width:100%">요청자료관리함</span></li>				
					<% else %>
								<li><span onClick="Ilban_read('')" style="width:100%">요청자료접수함</span></li>	
					<% END IF %>
				</ul>

			<% else %>

					<% if site_code <> "" then %>
						<h2>요청자료</h2>
						
						<% if site_code = "감리본부" then %>

							<li evt="0"><span onClick="Ilban_write('')" style="width:100%">요청자료발송함</span></li>	

						<% else %>
							
							<ul>
								<li evt="0"><span onClick="Ilban_read('')" style="width:100%">요청자료접수함</span></li>	
							</ul>
						<% end if %>
					<% END IF %>

			<% END IF %>

	<% END IF %>

<% 'END IF %>



	<% if db_id = "216050" or db_id = "208097" or db_id = "206171" or db_id = "205021" then %>
		<h2>공문결재현황</h2>
		<ul>			
			<% if UCASE(db_level) = "Z" or UCASE(db_level) = "S" THEN %>				
				<li><span onClick="appr_sheet('all')" style="width:100%">PM 결재현황</span></li>				
			<% elseif UCASE(db_level) = "P" then %>
				<li><span onClick="appr_sheet('mine')" style="width:100%">공문결재현황</span></li>	
			<% END IF %>
		</ul>
	<% end if %>

</div>







<script type="text/javascript">
	initToggleList(document.getElementById("left"), "h2", "ul", "li");

	initToggleList1(document.getElementById("left"), "h2", "ul", "li");


var currentListNum;
var level1El;
var level2El;
var level3El;
function initToggleList1(ulEl, level1, level2, level3)
{


	//alert("<%=Qang%>");
	//alert(ulEl);
	//alert(level1);
	//alert(level2);
	//alert(level3);

    currentListNum = true;
    
    level1El = ulEl.getElementsByTagName(level1);
    level2El = ulEl.getElementsByTagName(level2);
    level3El = ulEl.getElementsByTagName(level3);
    
   	for( var i = 0 ; i < level1El.length ; i++ )
   	{
		//alert("1");
		level1El.item(i).listNum = i;
		level1El.item(i).onclick = toggleList;
	}

	for( var i = 0 ; i < level2El.length ; i++ )
    {
		//alert("2");
		level2El.item(i).listNum = i;
		level2El.item(i).className = "off";
		level2El.item(i).subtag = level3;
	}
	
	for( var j = 0 ; j < level3El.length ; j++ )
  	{
		//alert("3");
  	    level3El.item(j).listNum = i;
  	    level3El.item(j).className = "off";
		level3El.item(j).onclick = toggleList_Sub;
		level3El.item(j).onmouseover = mouseOver_Sub;
		level3El.item(j).onmouseout = mouseOut_Sub;
	}
	
    if(level1 == "" && level3El.length > 0)
    {
		//alert("4");
		level3El.item(0).className = "on";
		prevSelMenu = level3El.item(0);
    }
	else if(level2El.length > 0){
		//alert("5");
		level2El.item(<%=Qang%>).className = "on";
		//level3El.item(1).onclick = toggleList_Sub;
	}
	else if(level1El.length > 1){
		//alert("6");
		level1El.item(0).className = "on";
	}

}
















</script>
</body>
</HTML>
