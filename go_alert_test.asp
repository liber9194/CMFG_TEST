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

<script language="javascript">
function goAlert()
{
	var f = document.form1;
	f.action = "http://192.168.32.238:12555";
	f.target = "custom";
	f.submit();
}
</script>

<body>
<!--<form method=post action="http://192.168.32.238:12555">-->
<form name="form1" method=post action="javascript:goAlert();">
	<input type="text" name="CMD" value="ALERT">
	<input type="text" name="Action" value="ALERT">
	<input type="text" name="Key" value= "alert_20170801_07">
	<input type="text" name="SystemName" value= "SFG">
	<input type="text" name="SystemName_Encode" value= "KSC5601">
	<input type="text" name="SendID" value= "216050">
	<input type="text" name="SendName" value= "류지현">
	<input type="text" name="SendName_Encode" value= "KSC5601">
	<input type="text" name="RecvID" value= "216050">
	<input type="text" name="Subject" value= "SFG 요청자료발송함 테스트">
	<input type="text" name="Subject_Encode" value= "KSC5601" >
	<textarea name="Contents">SFG 요청자료발송함 테스트 내용!</textarea>
	<input type="text" name="Contents_Encode" value= "KSC5601" >
	<!--<input type="text" name="URL" value= “http://www.ucware.net/aa.jsp?userid=(%USERID%)&password=(%USERPWD%)">-->
	<!--<input type="text" name="Option" value= "UM=POST,EN=BASE64H,SP=F,LT=20,TP=50,WD=600,HG=520,IT=N,IR=Y">-->
	<input type="submit" value="알림테스트[등록]">
	<!--<iframe name="custom" src="EmptyForAlarm.asp" frameborder="0" width="0" height="0"></iframe>-->
	<iframe name="custom" src="Empty.asp" frameborder="0" width="0" height="0"></iframe>
</form>


</body>
</html>

