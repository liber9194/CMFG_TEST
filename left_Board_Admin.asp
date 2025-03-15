<!-- #include file="../Mis/WebWrite/config.asp" -->
<!--#include file="../../../../dbopen.asp"-->
<!--#include file="../../../../ReqVariant.asp"-->
<!--#include file="../ezDoHwaBoard/misop.inc" -->
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
	function goPage(idx)
	{		
		try{
				var url = "";
				switch(idx)
				{
					case 1:
						url = "Right_Admin_Site.asp";
						break;
					
					case 2:
						url = "Right_Admin_GamriTuipList.asp";
						break;
						
					case 3:
						url = "Right_Admin_SuGum.asp";
						break;
						
					case 4:
						url = "Right_Admin_PM.asp" ;
						break;
						
					case 5:
						url = "Right_Admin_Mool.asp" ;
						break;
						
					case 6:
						url = "Right_Admin_KBN.asp" ;
						break;
						
					case 7:
						url = "Right_Admin_DongJel.asp" ;
						break;
						
					case 8:
						url = "Right_Admin_Tel.asp" ;
						break;
					case 9:
						url = "Right_Admin_GongMoon.asp";
						break;
					case 10:
						url = "Right_Admin_Dojang.asp";
						break;
					case 11:
						url = "Right_Admin_GongsaProgressRate.asp";
						break;		
					case 12:
						url = "Right_Admin_Include_PRJ.asp";
						break;	
					case 13:
						url = "Right_Admin_Set_Edu.asp";
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
		<div class="left_admin" title="관리자" style="width:100%">GFM(CM현장관리)</div>
			<h2><span onClick="goPage(1)">1.현장게시현황 검색</span><ul></ul></h2>	
			<h2><span onClick="goPage(2)">2.건설사업관리기술자<br>&nbsp;&nbsp;검색</span><ul></ul></h2>	
			<h2><span onClick="goPage(3)">3.CM 수금 계획 검색</span><ul></ul></h2>	
			<h2><span onClick="goPage(4)">4.CMFG PM 관리</span><ul></ul></h2>	
			<h2><span onClick="goPage(5)">5.CM 물가 변동 검색</span><ul></ul></h2>	
			<h2><span onClick="goPage(6)">6.현장게시현황 잠금(해제)</span><ul></ul></h2>	
			<h2><span onClick="goPage(7)">7.건설사업관리기술자<br>&nbsp;&nbsp;동절기 관리</span><ul></ul></h2>	
			<h2><span onClick="goPage(8)">8.현장연락처 보기</span><ul></ul></h2>	
			<h2><span onClick="goPage(9)">9.공문관리</span><ul></ul></h2>	
			<h2><span onClick="goPage(10)">10.사용인감</span><ul></ul></h2>	
			<h2><span onClick="goPage(11)">11.공사진도율 검색</span><ul></ul></h2>			
			<h2><span onClick="goPage(12)">12.준공(+1M)용역관리</span><ul></ul></h2>			
			<h2><span onClick="goPage(13)">13.비대면 온라인 교육</span><ul></ul></h2>				
	</div>

<script type="text/javascript">
	initToggleList(document.getElementById("left"), "h2", "ul", "li");
</script> 



</body>
</html>

