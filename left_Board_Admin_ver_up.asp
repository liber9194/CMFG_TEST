<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=11">	

<!-- #include file="../Mis/WebWrite/config.asp" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
<%
db_id 	 	= session("db_id")
db_level 	 	= session("db_level")

site_code 	 	= session("site_code")
%>

<link rel="stylesheet" href="../../Home/css/email_tree.css" type="text/css">
<link rel="stylesheet" href="../../Home/css/default_ver_up.css" type="text/css">
<link href="../../Home/skin/skin_1/skin.css" rel="stylesheet" type="text/css">
<style>
span
{
	font-size: 10pt;
	letter-spacing:-0.8px;
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="../../Home/myoffice/common/mouseeffect.js"></script>
<script language="javascript">
	function goPage(idx)
	{		
		try{
				var url = "";
				switch(idx)
				{
					case 0:
						url = "Right_Admin_Site_ver_up.asp";		
						break;
						
					case 1:
						url = "Right_Admin_Appointment_Check_List.asp";
						break;
					
					case 2:
						url = "Right_Admin_GamriTuipList_ver_up.asp";
						break;
						
					case 3:
						url = "Right_Admin_SuGum_ver_up.asp";
						break;
						
					case 4:
						url = "Right_Admin_PM_ver_up.asp" ;
						break;
						
					case 5:
						url = "Right_Admin_Mool_ver_up.asp" ;
						break;
						
					case 6:
						url = "Right_Admin_KBN_ver_up.asp" ;
						break;
						
					case 7:
						url = "Right_Admin_DongJel_ver_up.asp" ;
						break;
						
					case 8:
						url = "Right_Admin_Tel_ver_up.asp" ;
						break;
					case 9:
						url = "Right_Admin_GongMoon_ver_up.asp";
						break;
					case 10:
						url = "Right_Admin_Dojang_ver_up.asp";
						break;
					case 11:
						url = "Right_Admin_GongsaProgressRate_ver_up.asp";
						break;		
					case 12:
						url = "Right_Admin_Include_PRJ_ver_up.asp";
						break;	
					case 13:
						url = "Right_Admin_Set_Edu_ver_up.asp";
						break;	
					case 14:
						url = "Right_Admin_Notice_Popup.asp";
						break;
					case 15:
						// 직접 순서 수정, 조직 이름 변경, 조직 tree 설정 등...
						url = "Right_Admin_Jojik_Mng.asp";
						break;	
					case 16:
						// 직접경비정산>숙소 계약서 및 입금증 등록 화면 조회 (24.06.05.)
						url = "Right_Admin_DirectCost_House.asp";
						break;	
					case 17:
						// 기술지원 및 공통업무 출장비 비교 (24.12.16.)
						url = "Right_Admin_Tech_Common.asp";
						break;	
				}			
				window.open(url,"right");
			}
			catch(e){
				alert(e);
			}
	}
</script>

<body class="leftbody" leftmargin="0" topmargin="0" rightmargin="0" marginwidth="0" marginheight="0" style="overflow-y:auto;">

	<div id="left">
		<div class="left_admin" title="CMFG 관리자 메뉴"></div>
			<h2><span onClick="goPage(0)">현장게시현황 검색</span></h2><ul></ul> 		
			<h2><span onClick="goPage(1)">지킴약속 Check List</span></h2><ul></ul> 		
			<h2><span onClick="goPage(17)">기술지원 및 공통업무<br>출장비 비교</span></h2><ul></ul> 		
			<h2><span onClick="goPage(16)">숙소(계약서,입금증) 현황</span></h2><ul></ul>	
			<h2><span onClick="goPage(2)">건설사업관리기술자 검색</span></h2><ul></ul>	
			<h2><span onClick="goPage(3)">CM 수금 계획 검색</span></h2><ul></ul>	
			<h2><span onClick="goPage(4)">CMFG PM 관리</span></h2><ul></ul>	
			<h2><span onClick="goPage(5)">CM 물가 변동 검색</span></h2><ul></ul>
			<h2><span onClick="goPage(6)">현장게시현황 잠금(해제)</span></h2><ul></ul>	
			<h2><span onClick="goPage(7)">건설사업관리기술자<br>동절기 관리</span></h2><ul></ul>	
			<h2><span onClick="goPage(8)">현장연락처 보기</span></h2><ul></ul>	
			<h2><span onClick="goPage(9)">공문관리</span></h2><ul></ul>	
			<h2><span onClick="goPage(10)">사용인감</span></h2><ul></ul>	
			<h2><span onClick="goPage(11)">공사진도율 검색</span></h2><ul></ul>			
			<h2><span onClick="goPage(12)">준공(+1M)용역관리</span></h2><ul></ul>			
			<h2><span onClick="goPage(13)">비대면 온라인 교육</span></h2><ul></ul>	
			<h2><span onClick="goPage(14)">공지 관리</span></h2><ul></ul>	
			<h2><span onClick="goPage(15)">CM 조직도관리</span></h2><ul></ul>	
	</div>

<script type="text/javascript">
	initToggleList(document.getElementById("left"), "h2", "ul", "li");	
</script> 

</body>
</html>