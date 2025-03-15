<%@ LANGUAGE="VBSCRIPT" %>
<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=11">	

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

<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="../../Home/css/email_tree.css" type="text/css">
<!--<link rel="stylesheet" href="../../Home/css/default.css" type="text/css">-->
<link rel="stylesheet" href="../../Home/css/default_ver_up.css" type="text/css">
<link href="../../Home/skin/skin_1/skin.css" rel="stylesheet" type="text/css">
<style>
span
{
	width:100%;
	letter-spacing:-0.8px;
}
</style>

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

/*
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
		
		window.parent.frames("right").location.href = "/myoffice/ezBoardSTD/BoardItemList.aspx?BoardID=" + RedirectBoardID;
	}
}
*/

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
	xmlDom_treeview.load("/myoffice/ezBoardSTD/config/BoardTree_config.xml");
	
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
		
		window.parent.frames(2).location.href = "/myoffice/ezBoardSTD/BoardItemList.aspx?BoardID=" + SelectedTreeView.getvalue(nodeIdx, "DATA1") + "&BoardName=" + SelectedTreeView.getvalue(nodeIdx, "DATA2");
		
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
	xmlhttp.open("POST", "/myoffice/ezBoardSTD/interASP/GetSubBoards.aspx?RootBoardID=top&SubFlag=0", false);
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
	xmlDom_treeview.load("/myoffice/ezBoardSTD/config/BoardTree_config.xml");
	
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
	xmlhttp4.open("POST", "/myoffice/ezBoardSTD/interASP/GetMyBoards.aspx", false);
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
	xmlDom_treeview.load("/myoffice/ezBoardSTD/config/BoardTree_config.xml");
	
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
	xmlhttp3.open("POST", "/myoffice/ezBoardSTD/interASP/GetSubBoards.aspx?RootBoardID=" + pRootBoardID + "&SubFlag=" + pSubFlag + "&SelectFlag=0", false);
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
		strHTML += "	  <div  class='tree' id='TreeCtrl" + i.toString() + "' style='behavior:url(/myoffice/ezBoardSTD/Controls/BoardTreeview.htc);height:100%;width:100%;overflow-x:auto;overflow-y:auto;padding-left:10px' onreadystatechange='TreeCtrl_onreadystatechange();' onrequestdata='TreeCtrl_onNodeExpanded();' onnodeselect='TreeCtrl_onNodeClick();TreeCtrl" + i.toString() + ".toggle(TreeCtrl" + i.toString() + ".selectedIndex)'></div>";
		strHTML += "  </ul>";
	}
	
	xmldomNodes = null;
	xmldom = null;
	
	TopBoardsList.innerHTML = strHTML;
}

function AdminMenu_onclick()
{
	window.open("/myoffice/ezBoardSTD/admin/index_admin.aspx", "", "height=" + window.screen.availHeight + ",width=" + window.screen.availWidth + ", status = no, toolbar=no, menubar=no, location=no, resizable=1, left=0, top=0","");	
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
		xmlhttp5.open("POST", "/myoffice/ezBoardSTD/interASP/DeleteMyBoard.aspx?BoardID=" + TreeCtrl_MyBoardTree.getvalue(nodeIdx, "DATA1"), false);
		xmlhttp5.send();
		xmlhttp5 = null;
		TreeCtrl_MyBoardTree.source = GetMyBoardItem();
		TreeCtrl_MyBoardTree.update();
	}
}

function Open_Func(idx)
{
	if(idx== 1)
		window.parent.frames("right").location.href	= "/myoffice/ezQuestion/poll/Qst_List.aspx?brd_ID=5";
	else
		window.parent.frames("right").location.href	= "/myoffice/ezQuestion/poll/Qst_Step1.aspx?brd_ID=5"

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

function D_fUN(Pcode,Pstr)
{
	parent.frames[2].location = "Right_HSBoard_ver_up.asp?code=" + Pcode + "&Qstr=" + Pstr;
}

function D_Chk_Disaster(Pstr)
{
	parent.frames[2].location = "Right_Chk_Disaster_ver_up.asp?Qstr=" + Pstr;
}

</script>

<body class="leftbody" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="overflow-y:auto;"> 
<!-- PDS_TBL 테이블의 dk_code 8xx 는 안전보건관리 영역 -->
<div id="left">
		<div class="left_HSboard" title="안전·보건관리"></div>
		
		<h2>안전·보건 경영메뉴얼</h2>
		<ul>		     
					<li><span style="width:100%" onClick="D_fUN('800','본사 안전·보건 경영메뉴얼')">본사 안전·보건 경영메뉴얼</span></li>
					<li><span style="width:100%" onClick="D_fUN('801','본사 안전·보건 경영메뉴얼')">현장 안전·보건 경영메뉴얼</span></li>
					<li><span style="width:100%" onClick="D_fUN('802','안전·보건경영방침')">안전·보건경영방침</span></li>
		</ul> 		
		
		<h2>안전·보건 관련법</h2>
		<ul>		     
					<li><span style="width:100%" onClick="D_fUN('820','안전보건 관련 법령')">안전보건 관련 법령</span></li>
					<li><span style="width:100%" onClick="D_fUN('821','안전보건 관련 지침')">안전보건 관련 지침</span></li>
					<li><span style="width:100%" onClick="D_fUN('822','안전보건 관련 보도자료')">안전보건 관련 보도자료</span></li>
					<li><span style="width:100%" onClick="D_fUN('823','안전보건 관련 행정규칙')">안전보건 관련 행정규칙</span></li>
		</ul>
		
		<h2>안전관리 수준평가</h2>
		<ul>		     
					<li><span style="width:100%" onClick="D_fUN('840','본사 안전관리 수준평가')">본사 안전관리 수준평가</span></li>
					<li><span style="width:100%" onClick="D_fUN('841','현장 안전관리 수준평가')">현장 안전관리 수준평가</span></li>
					<li><span style="width:100%" onClick="D_fUN('842','연도별 안전관리수준평가 현황')">연도별 안전관리수준평가<br>&nbsp;&nbsp;&nbsp;&nbsp;현황</span></li>
		</ul>

		<h2>안전사고 보고 및 사례</h2>
		<ul>		     
					<li><span style="width:100%" onClick="D_fUN('860','안전사고 보고')">안전사고 보고</span></li>
					<li><span style="width:100%" onClick="D_fUN('861','안전사고 사례')">안전사고 사례</span></li>
		</ul> 


		<h2>안전·보건 교육 자료</h2>
		<ul>		     
					<li><span style="width:100%" onClick="D_fUN('870','안전교육 자료')">안전교육 자료</span></li>
					<li><span style="width:100%" onClick="D_fUN('871','보건교육 자료')">보건교육 자료</span></li>
		</ul> 


		<h2>본사안전팀 회의 자료</h2>
		<ul>		     
					<li><span style="width:100%" onClick="D_fUN('880','본사안전팀 회의 자료')">본사안전팀 회의 자료</span></li>
		</ul> 	

		<h2>기타 안전 관련 자료</h2>
		<ul>		     
					<li><span style="width:100%" onClick="D_fUN('890','기타 안전 관련 자료')">기타 안전 관련 자료</span></li>								
		</ul> 		

		<!-- 중대재해처벌법 이행사항 메뉴 추가 (22.03.29 남승혜 사원 요청) -->
		<h2>중대재해처벌법 이행사항</h2>
		<ul>	
				<li><span style="width:100%" onClick="D_Chk_Disaster('중대재해처벌법 이행사항')">중대재해처벌법 이행사항</span></li>
		</ul>
		

			
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
debugger;
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