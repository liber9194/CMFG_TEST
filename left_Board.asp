<!-- #include file="../Mis/WebWrite/config.asp" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<!-- #include file="../../../../default_properties.asp" -->
<%


db_id 	 	= session("db_id")
db_level 	 	= session("db_level")

site_code 	 	= session("site_code")

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<html>
<head>
<link rel="stylesheet" href="../../Home/css/email_tree.css" type="text/css">
<link rel="stylesheet" href="../../Home/css/default.css" type="text/css">
<link href="../../Home/skin/skin_1/skin.css" rel="stylesheet" type="text/css">

<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">

<script language="JavaScript" src="../../Home/myoffice/common/mouseeffect.js"></script>
<script language="javascript">
var SSUserID = "204112";
var SSUserName = "이재훈";
var SSDeptID = "180___";
var SSDeptName = "정보기술실";
var SSCompanyID = "dohwa";
var SSCompanyName = "도화종합기술공사";

var SelectedBoardID = "";
var SelectedBoardParentBoardID = "";

var SS_ServerName = "gw.dohwa.co.kr";

var RedirectBoardGroupID = "{7AF62F0E-3E0B-4639-9C53-CA0B355546E5}";
var RedirectBoardID = "{63583DD8-B4AF-4689-BAF1-26216953A5EA}";

var Func = "";

//[070530]_마이게시판
var PhotoType = "";
var g_ReadyState = "";

function window.onload()
{
	if (Func == "1")
	{
		WebPartToggle(level1El.item(level1El.length-1));
		
		Open_Func(1);
	}
	else if (RedirectBoardID == "" || RedirectBoardGroupID == "")
	{
		ShowMyBoardItem();	
	}
	
	if( RedirectBoardID != "" )
	{
	    	//ShowMyBoardItem();
		document.all("{00000000-0000-0000-0000-000000000000}").parentElement.click();
		if (RedirectBoardGroupID != "" && g_ReadyState == "") BoardRedirect();
		
		window.parent.frames("right").location.href = "../ezBoardSTD/BoardItemList.aspx?BoardID=" + RedirectBoardID;
	}
}

function BoardRedirect()
{
	var spans = TopBoardsList.all.tags("span");
	
	for( var i = 0 ; i < spans.length ; i++ )
	{
		if( spans.item(i).id == RedirectBoardGroupID )
		{
			LoadTreeViewByPath(spans.item(i), RedirectBoardID, RedirectBoardGroupID);
		}
	}
}

function LoadTreeViewByPath(pObjSpan, pBoardID, pBoardGroupID)
{
	pObjSpan.parentElement.click();
	
	var TreeCtrl = pObjSpan.parentElement.nextSibling.firstChild.firstChild;
	
	var xmlDom_treeview = new ActiveXObject("Microsoft.XMLDOM");
	xmlDom_treeview.async = false;
	xmlDom_treeview.load("../ezBoardSTD/config/BoardTree_config.xml");
	
	TreeCtrl.server = SS_ServerName;
	TreeCtrl.config = xmlDom_treeview;
	TreeCtrl.source = GetBoardTreeByPath(pBoardID, pBoardGroupID);
	TreeCtrl.update();
	
	xmlDom_treeview = null;
}

function GetBoardTreeByPath(pBoardID, pBoardGroupID)
{
	var xmlhttp2 = new ActiveXObject("Microsoft.XMLHTTP");
	xmlhttp2.open("POST", "admin/interASP/GetBoardTreeByPath.aspx?BoardID=" + pBoardID + "&BoardGroupID=" + pBoardGroupID, false);
	xmlhttp2.send();

	var ret = xmlhttp2.responseXML;
	xmlhttp2 = null;

	return ret;
}

function TreeCtrl_onNodeExpanded() 
{
	var SelectedTreeView = window.event.srcElement;
	var nodeIdx = window.event.nodeIdx;

	if(SelectedTreeView.id == "TreeCtrl_MyBoardTree") return;

	var childXML = GetSubBoard(SelectedTreeView.getvalue(nodeIdx, "DATA1"), "1")
	SelectedTreeView.putchildxml(nodeIdx, childXML);
}

function TreeCtrl_onNodeClick() 
{	
	try 
	{
		var SelectedTreeView = window.event.srcElement;
		var nodeIdx = SelectedTreeView.selectedIndex;
		
		//20080130_성수곤
		//현재 선택한 그룹이외의 게시판트리뷰는 모두 Unselect 한다.
		SetTreeviewUnSelect(SelectedTreeView.id);

		SelectedBoardID = SelectedTreeView.getvalue(nodeIdx, "DATA1");
		SelectedBoardParentBoardID = SelectedTreeView.getvalue(nodeIdx, "DATA3");
		var chkPhotoBrd = SelectedTreeView.getvalue(nodeIdx, "DATA5");	// 20070228 포토게시판 추가
		
		window.parent.frames(2).location.href = "../ezBoardSTD/BoardItemList.aspx?BoardID=" + SelectedTreeView.getvalue(nodeIdx, "DATA1") + "&BoardName=" + SelectedTreeView.getvalue(nodeIdx, "DATA2");
		
		window.event.cancelBubble = true;
		window.event.returnValue = false;
	}
	catch(e)
	{
		alert(e.description);
	}
}

//20081030_성수곤
//파라미터의 트리뷰만을 제외하고 나머지 트리뷰는 모두 unselect
function SetTreeviewUnSelect(TreeviewID)
{
	for(i=0;i<document.all.tags("div").length;i++)
	{
		if(document.all.tags("div")(i).id.indexOf('TreeCtrl') > -1 && document.all.tags("div")(i).id != TreeviewID)
		{
			document.all.tags("div")(i).unselect();
		}
	}
}

function DisplayTopBoard()
{
	var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
	xmlhttp.open("POST", "../ezBoardSTD/interASP/GetSubBoards.aspx?RootBoardID=top&SubFlag=0", false);
	xmlhttp.send();
	
	if(xmlhttp.responseXML.text != "ERROR")
	{
		MakeTopBoardView(xmlhttp.responseXML.xml)
	}

	xmlhttp = null;
}

function ShowMyBoardItem()
{
	var xmlDom_treeview = new ActiveXObject("Microsoft.XMLDOM");
	xmlDom_treeview.async = false;
	xmlDom_treeview.load("../ezBoardSTD/config/BoardTree_config.xml");
	
	TreeCtrl_MyBoardTree.server = SS_ServerName;
	TreeCtrl_MyBoardTree.config = xmlDom_treeview;
	TreeCtrl_MyBoardTree.source = GetMyBoardItem();
	try{
	    TreeCtrl_MyBoardTree.update();	
	}catch(e){}
	
	xmlDom_treeview = null;
}

function GetMyBoardItem()
{
	var xmlhttp4 = new ActiveXObject("Microsoft.XMLHTTP");
	xmlhttp4.open("POST", "../ezBoardSTD/interASP/GetMyBoards.aspx", false);
	xmlhttp4.send();
	var ret = xmlhttp4.responseXML;
	xmlhttp4 = null;
	return ret;
}

function TopBoard_onclick(obj)
{
	if (g_ReadyState != "") return;
	
	var TopBoardID = window.event.srcElement.id;	
	var TreeCtrl = obj;
	var xmlDom_treeview = new ActiveXObject("Microsoft.XMLDOM");
	xmlDom_treeview.async = false;
	xmlDom_treeview.load("../ezBoardSTD/config/BoardTree_config.xml");
	
	TreeCtrl.server = SS_ServerName;
	TreeCtrl.config = xmlDom_treeview;
	TreeCtrl.source = GetSubBoard(TopBoardID, "0");
	TreeCtrl.update();
	
	xmlDom_treeview = null;
}

function TreeCtrl_onreadystatechange()
{
	if (event.srcElement.readyState == "loading") 
	{
		if (g_ReadyState.indexOf(event.srcElement.id + ";") == -1)
			g_ReadyState += event.srcElement.id + ";";
	}
	else if (event.srcElement.readyState == "complete") 
	{
		g_ReadyState = g_ReadyState.replace(event.srcElement.id + ";", "");
	}
}

function GetSubBoard(pRootBoardID, pSubFlag)
{
	var xmlhttp3 = new ActiveXObject("Microsoft.XMLHTTP");
	xmlhttp3.open("POST", "../ezBoardSTD/interASP/GetSubBoards.aspx?RootBoardID=" + pRootBoardID + "&SubFlag=" + pSubFlag + "&SelectFlag=0", false);
	xmlhttp3.send();	

	var ret = xmlhttp3.responseXML;	

	xmlhttp3 = null;
	return ret;
}

function MakeTopBoardView(strXML)
{
	var xmldom = new ActiveXObject("Microsoft.XMLDOM");
	var strHTML = "";
	xmldom.async = false;
	xmldom.preserveWhiteSpace = true;
	xmldom.loadXML(strXML);
	
	strHTML = "";
	var xmldomNodes = xmldom.selectNodes("TREEVIEWDATA/NODE");
	for(i=0;i<xmldomNodes.length;i++)
	{
		strHTML += "<h2><span id='" + xmldomNodes.item(i).selectSingleNode("DATA1").text + "' onclick='TopBoard_onclick(\"TreeCtrl"+i.toString()+"\")'>" + xmldomNodes.item(i).selectSingleNode("DATA2").text + "</span></h2>";
		strHTML += "  <ul>";
		strHTML += "	  <div  class='tree' id='TreeCtrl" + i.toString() + "' style='behavior:url(../ezBoardSTD/Controls/BoardTreeview.htc);height:100%;width:100%;overflow-x:auto;overflow-y:auto;padding-left:10px' onreadystatechange='TreeCtrl_onreadystatechange();' onrequestdata='TreeCtrl_onNodeExpanded();' onnodeselect='TreeCtrl_onNodeClick();TreeCtrl" + i.toString() + ".toggle(TreeCtrl" + i.toString() + ".selectedIndex)'></div>";
		strHTML += "  </ul>";
	}
	
	xmldomNodes = null;
	xmldom = null;
	
	TopBoardsList.innerHTML = strHTML;
}

function AdminMenu_onclick()
{
	window.open("../ezBoardSTD/admin/index_admin.aspx", "", "height=" + window.screen.availHeight + ",width=" + window.screen.availWidth + ", status = no, toolbar=no, menubar=no, location=no, resizable=1, left=0, top=0","");	
}

function DeleteMyBoard()
{
	var nodeIdx = TreeCtrl_MyBoardTree.selectedIndex;
	if(TreeCtrl_MyBoardTree.getvalue(nodeIdx, "DATA1") == "{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}") {
		alert("새 게시 게시판은 삭제할 수 없습니다.");
		return;
	}
	var ret = confirm(TreeCtrl_MyBoardTree.getvalue(nodeIdx, "DATA2") + " 게시판을 마이게시판에서 삭제하시겠습니까?");
	if(ret) {
		var xmlhttp5 = new ActiveXObject("Microsoft.XMLHTTP");
		xmlhttp5.open("POST", "../ezBoardSTD/interASP/DeleteMyBoard.aspx?BoardID=" + TreeCtrl_MyBoardTree.getvalue(nodeIdx, "DATA1"), false);
		xmlhttp5.send();
		xmlhttp5 = null;
		TreeCtrl_MyBoardTree.source = GetMyBoardItem();
		TreeCtrl_MyBoardTree.update();
	}
}

function Open_Func(idx)
{
	if(idx== 1)
		window.parent.frames("right").location.href	= "../ezQuestion/poll/Qst_List.aspx?brd_ID=5";
	else
		window.parent.frames("right").location.href	= "../ezQuestion/poll/Qst_Step1.aspx?brd_ID=5"

	//20080130_성수곤
	//현재 선택한 그룹이외의 게시판트리뷰는 모두 Unselect 한다.
	SetTreeviewUnSelect("");
}

// 신규 토글 함수
function WebPartToggle(obj)
{
	//level1El.item(0).className = "off";
	//level2El.item(0).className = "off";		 
	
	if( obj.listNum && currentListNum != obj.listNum +1)
	{
		level1El.item(currentListNum -1).className = null;
		level2El.item(currentListNum -1).className = "off";
	}
	
	if(level2El.item(obj.listNum).className == "on" )
	{
		level1El.item(obj.listNum).className = null;
		level2El.item(obj.listNum).className = "off";
	}
	else
	{
		level1El.item(obj.listNum).className ="on";
		level2El.item(obj.listNum).className ="on";
	}
	
	currentListNum = obj.listNum + 1;
	
	setMenu(level2El.item(obj.listNum)); 
}



	function write_Letter()
	{
		parent.frames[2].location = "Right_Site.asp";
	}
	function go_article()
	{
		parent.frames[2].location = "Right_Article_Main.asp";
	}


	function write_Letter_PM()
	{
		parent.frames[2].location = "Right_Site_PM.asp";
	}

	function write_Letter_NOT()
	{
		alter("현장 직원이 아니므로 진행할수 없습니다");
	}


	function write_Letter1()
	{
		parent.frames[2].location = "super_board.asp";
	}






	function write_Letter_pppp()
	{
		parent.frames[2].location = "Right_Address_New.asp?st=1&c_M=0";
	}



	function write_Letter_pppp_01()
	{
		parent.frames[2].location = "Right_Address_New_01.asp?st=1&c_M=0";
	}


	//2016.04.25 주석처리
	//function write_Letter_pppp1()
	//{
	//	parent.frames[2].location = "Right_Address_New.asp?st=1&c_M=1";
	//}
	
	// 기관감사 수감 현황
	// 2015-02-24 추가, 2016.04.25 주석처리
	//function write_Letter_org_Supervising()
	//{
	//	parent.frames[2].location = "Right_Org_Supervising.asp";
	//}



</script>

<body class="leftbody" leftmargin="0" topmargin="0" rightmargin="0" marginwidth="0" marginheight="0" style="overflow-y:auto;">
	<div id="left">
		<div class="left_board" title="PIMS" style="width:100%"></div>
			<% if db_level = "P" OR db_level = "Z" OR db_level = "S" THEN %>
				<h2><span id="{00000000-0000-0000-0000-000000000000}" onClick="write_Letter()"  style='width:100%'>현장게시현황</span></h2>
					<ul>					
					</ul> 
			<% else %>
				<% if site_code = "" THEN %>				
					<h2><span id="{00000000-0000-0000-0000-000000000000}" onClick="write_Letter_NOT()"  style='width:100%'>현장게시현황</span></h2>			
					<ul>
					</ul> 					
				<% else %>
					<h2><span id="{00000000-0000-0000-0000-000000000000}" onClick="write_Letter()"  style='width:100%'>현장게시현황</span></h2>		
					<ul>
					</ul> 					
				<% end if %>
			<% end if %>


			<!--h2>도움서(Q/A)</h2-->
			<h2>현장게시현황 도움서</h2>
			<ul>		     
					<!--<li><a href="http://sfg.dohwa.co.kr/data/menu/도화SFG사용설명서.hwp" target="_blank" id="white">도화CMFG사용설명서</a> </li>-->
					<li><a href="<%=g_file_path%>/data/menu/도화SFG사용설명서.hwp" target="_blank" id="white">도화CMFG사용설명서</a> </li>
					<li><a href="<%=g_file_path%>/data/menu/도화SFG현장게시현황.hwp" target="_blank" id="white">도화CMFG현장게시현황</a></li>
					<li><a href="<%=g_file_path%>/data/menu/주간업무및공정보고양식.hwp" target="_blank" id="white">주간업무및공정보고양식</a></li>
					<li><a href="<%=g_file_path%>/data/menu/대표이사-발주청(공문).hwp" target="_blank" id="white">대표이사-발주청<br>&nbsp;&nbsp;&nbsp;&nbsp;(공문양식)</a></li>
					<li><a href="<%=g_file_path%>/data/menu/책임감리원-대표이사(공문).hwp" target="_blank" id="white">건설사업관리단-대표이사(공문양식)</a></li>
			</ul> 	
			<!-- 주석 Start
				<ul>
					<div class="tree" style='behavior:url(../ezBoardSTD/Controls/BoardTreeview.htc);height:100%;width:100%;overflow-x:hidden;margin-left:5px' id='TreeCtrl_MyBoardTree' onreadystatechange='TreeCtrl_onreadystatechange();' onrequestdata='TreeCtrl_onNodeExpanded();' onnodeselect='TreeCtrl_onNodeClick();TreeCtrl_MyBoardTree.toggle(TreeCtrl_MyBoardTree.selectedIndex)'></div><br>
					<li evt="0"><span onClick="DeleteMyBoard()" style="width:100%">마이게시판삭제</span></li>
				</ul> 
			<div id='TopBoardsList'>
			<h2><span style='width:100%' id='{7AF62F0E-3E0B-4639-9C53-CA0B355546E5}' onclick='TopBoard_onclick(TreeCtrl0)'>현장개요(요약)</span></h2>

				<ul>
					<div class="tree" style='behavior:url(../ezBoardSTD/Controls/BoardTreeview.htc);height:100%;width:100%;overflow-x:hidden;margin-left:5px' id='TreeCtrl_MyBoardTree' onreadystatechange='TreeCtrl_onreadystatechange();' onrequestdata='TreeCtrl_onNodeExpanded();' onnodeselect='TreeCtrl_onNodeClick();TreeCtrl_MyBoardTree.toggle(TreeCtrl_MyBoardTree.selectedIndex)'></div><br>
					<li evt="0"><span onClick="DeleteMyBoard()" style="width:100%">마이게시판삭제</span></li>
				</ul> 
			 주석 End -->


			<h2><span style='width:100%' id='{E60DF53A-2738-4C7C-A2CF-F81CA4D37ED9}' onClick="write_Letter1()">조직도</span></h2>

			<% if db_level = "P" or db_level = "Z"  THEN %>
					<!-- 2016.04.25 주석처리
					<h2><span style='width:100%' id='{531B4347-C914-4EC8-B057-58A5F350E96A}' onclick='write_Letter_PM()'>PM 관리 현장</span></h2>
					-->
					<h2><span style='width:100%' id='{E60DF53A-2738-4C7C-A2CF-F81CA4D37ED9}' onClick="write_Letter_pppp()">CM단연락처</span></h2>
					<!--2016.04.22 주석 처리 -->
					<!--h2><span style='width:100%' id='{E60DF53A-2738-4C7C-A2CF-F81CA4D37ED9}' onClick="write_Letter_pppp1()">현장연락처보기_수정</span></h2-->
					<h2><span style='width:100%' id='{E60DF53A-2738-4C7C-A2CF-F81CA4D37ED9}' onClick="write_Letter_pppp_01()">CM부문(핸드폰)</span></h2>
			<% elseif db_level = "D" then %>
				<h2><span style='width:100%' id='{E60DF53A-2738-4C7C-A2CF-F81CA4D37ED9}' onClick="write_Letter_pppp()">CM단연락처</span></h2>
			<% end if %>
			
			<!-- 기관감사 수감 현황 (2015-02-24) 추가 -->
			<% IF db_id = "215053" or db_id = "204112" or db_id = "206171" or db_id = "213106" THEN %>
					<!-- 2016.04.05 주석처리
					<h2><span style='width:100%' id='{E60DF53A-2738-4C7C-A2CF-F81CA4D37ED9}' onClick="write_Letter_org_Supervising()">기관감사 수감 현황</span></h2>
					-->
			<% END IF %>
			<!--2016.04.20 추가 -->
			<h2><span style='width:100%' id='spanMenuArticle' onClick="go_article()">현장물품보유현황</span></h2>					

		<h2 style="display:none"><span id="DEPTCONT" onClick="Open_Func(this)" style="width:100%"></span></h2>
		<ul style="display:none">
		<div class="tree" id="DeptContTree" valign="top" style="behavior:url(/myoffice/common/organtreeview.htc);height:160px;width:100%;overflow-x:auto;overflow-y:auto;background-color:#FFFFFF;" 
			onrequestdata="DeptContRequestData()" onnodeselect="DeptContTree.toggle(DeptContTree.selectedIndex);DeptContNodeClick()" 
			onnodedblclick=""></div>
			
		</ul>
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
</html>

