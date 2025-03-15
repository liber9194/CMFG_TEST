<%@ Language=VBScript %>
<%
	'WebWrite_dir        = Request("WebWrite_dir")
	'WebWrite_content    = Request("WebWrite_content")
%>

<html>
<head>

		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ks_c_5601-1987">

    <title>SFG 저장</title>
    <!-- DEXTUploadX 버튼 상태를 위한 변수 -->

    <script language="javascript">
            // 전송중 상태를 위한 변수
            var g_bTransfer = false; 


			var Pcode = "";
			var PQstr = "";
			var PKtype = "";

    </script>

    <!-- DEXTUploadX 에러처리 -->

    <script for="FileUploadMonitor" event="OnError(nCode, sMsg, sDetailMsg)" language="javascript">
	       	OnFileMonitorError(nCode, sMsg, sDetailMsg);
    </script>

    <script language="javascript">
    	    function OnFileMonitorError(nCode, sMsg, sDetailMsg)
    	    {
    		    alert(nCode);
    		    alert(sMsg);
    		    alert(sDetailMsg);
    	    }
    </script>

    <!-- DEXTUploadX 전송 정보 초기화 -->

    <script language="javascript">
            function Init()
            {
                // 파일개수 
                document.getElementById('_CurrentCount').innerHTML = 0; 
                document.getElementById('_TotalCount').innerHTML = document.all["FileUploadMonitor"].Count; 
                                
                // 전체파일 전송량
                document.getElementById('_TransferedSizeOfTotalFile').innerHTML = "0%"; 
                                
                // 남은시간
                document.getElementById('_RemainingTime').innerHTML = "00:00:00"; 
                
                // 전체 파일 사이즈에 대한 프로그래스바
                document.all["Progress_total"].CurrentPos = 0; 
                
                // 파일이름
                document.getElementById('_CurrentFileName').innerHTML = ""; 
                
                // 현재파일 전송량
                document.getElementById('_TransferedSizeOfCurrentFile').innerHTML = "0%"; 
                
                // 전송속도
                document.getElementById('_BytesPerSec').innerHTML = "0 MB";  	
                
                // 현재 파일 사이즈에 대한 프로그래스바
                document.all["Progress_file"].CurrentPos = 0; 
                
                // 체크박스 초기화
                //if(true == document.all["FileUploadMonitor"].CheckAutoCloseWindow) 
			    //{
			    //    document.images["Chk_Box"].src="../../img/New_image/btn_checkyes.GIF"; 
			    //} 
			    //else
			    //{
			    //    document.images["Chk_Box"].src="../../img/New_image/btn_checkno.GIF"; 
			    //}
				//alert("end");
            }
    </script>

    <!-- DEXTUploadX 전송 정보 리프레쉬 -->

    <script language="javascript">
    	    function Refresh()
    	    {
        	    // 파일개수
    	        document.getElementById('_CurrentCount').innerHTML = document.all["FileUploadMonitor"].CurrentCount;  
    	        document.getElementById('_TotalCount').innerHTML = document.all["FileUploadMonitor"].Count; 
        	    
    	        // 전체파일 전송량
    	        document.getElementById('_TransferedSizeOfTotalFile').innerHTML = document.all["FileUploadMonitor"].PercentOfTotalFile + "%";
        	    
    	        // 남은시간 
    	        document.getElementById('_RemainingTime').innerHTML = document.all["FileUploadMonitor"].RemainingTime;  
        	    
    	        // 전체 파일의 사이즈에 대한 프로그래스바   
				var g_currnetPos_total = 0;
					// "FileUploadMonitor" 에서 값을 받아와 프로그래스를 지정
			    g_currnetPos_total = document.all["FileUploadMonitor"].PercentOfTotalFile; 
				 document.all["Progress_total"].CurrentPos = g_currnetPos_total; 

    	        // 파일이름(전체경로에서 파일명 분리, 파일명의 길이가 12자를 넘을 경우 "..." 처리)
    	        var fullpath = document.all["FileUploadMonitor"].CurrentFileName;
    	        var c_FileName = fullpath.substring(fullpath.lastIndexOf("\\") + 1);
    	        var name_Len = c_FileName.length;
    	        var cut_Name = "";
    	        
    	        if(55 < name_Len)
    	        {
    	            cut_Name = c_FileName.substr(0,55) + "...";
    	            document.getElementById('_CurrentFileName').innerHTML = cut_Name;
    	        }
    	        else
    	        {
    	            document.getElementById('_CurrentFileName').innerHTML = c_FileName;
    	        }
    	          	    
    	        // 현재파일 전송량
    	        document.getElementById('_TransferedSizeOfCurrentFile').innerHTML = document.all["FileUploadMonitor"].PercentOfCurrentFile + "%";
    	        
    	        // 전송속도
    	        var CurrentSpeed = document.all["FileUploadMonitor"].BytesPerSec; 

    	        if(1 < CurrentSpeed/1000.0/1000.0/1000.0) 
    	        {
    	           CurrentSpeed = CurrentSpeed / (1024*1024*1024);
    	           document.getElementById('_BytesPerSec').innerHTML = Math.round(CurrentSpeed * 100) / 100 + " GB";
    	        }
    	        else if(1 < CurrentSpeed/1000.0/1000.0) 
    	        {
    	           CurrentSpeed = CurrentSpeed / (1024*1024); 
    	           document.getElementById('_BytesPerSec').innerHTML = Math.round(CurrentSpeed * 100) / 100 + " MB";  
    	        }
    	        else if(1 < CurrentSpeed/1000.0) 
    	        {  
    	           CurrentSpeed = CurrentSpeed / 1024;
    	           document.getElementById('_BytesPerSec').innerHTML = Math.round(CurrentSpeed * 100) / 100 + " KB";
    	        }
    	        else 
    	        {
    	           document.getElementById('_BytesPerSec').innerHTML = Math.round(CurrentSpeed * 100) / 100 + " Bytes";
    	        }  	
        	    
    	        // 현재 파일의 사이즈에 대한 프로그래스바   
				var g_currnetPos_file = 0;
					// "FileUploadMonitor" 에서 값을 받아와 프로그래스를 지정
			    g_currnetPos_file = document.all["FileUploadMonitor"].PercentOfCurrentFile; 
				 document.all["Progress_file"].CurrentPos = g_currnetPos_file;
    	    }
    </script>

    <!-- DEXTUploadX 리프레쉬 타임 설정 -->

    <script language="JavaScript">
    
            var termination = 0;  
            var time; 
           
    	    function Repeat(bRefresh) 
    	    {    
    	        termination = document.all["FileUploadMonitor"].PercentOfTotalFile;
        	    
    	        // 전송 상태 리프레쉬
    	        Refresh(); 
                           
    	        if(bRefresh == false) 
                { 
    	           // 다운로드 취소
    	           clearTimeout(time); 
    	           Init() 
    	        }
    	        else  
                {  					

                    if (termination >= 100)
                    {


					   //alert("11");
                       // 다운로드 완료
                       //document.images["Cancel"].src="../../img/New_image/btn_close_nor.gif"; 
					   //alert("12");


					   window.opener.close()	
					   window.close()
                       clearTimeout(time); 
					   //alert("12");
                       //if(true != document.all["FileUploadMonitor"].CheckAutoCloseWindow) {

                           //alert("전송이 완료되었습니다.");  
						   //alert(document.all("FileUploadMonitor").ResponseData);

								
								//alert(PKtype);
								//alert(Pcode);
								//alert(PQstr);

								if (PKtype == '1'){
									window.opener.opener.location.href = "../../ezDoHwaBoard/Right_JaRyo_Upmubogo.asp?code=" + Pcode + "&Qstr=" + PQstr ;
								} else {
									window.opener.opener.location.href = "../../ezDoHwaBoard/Right_JaRyo.asp?code=" + Pcode + "&Qstr=" + PQstr ;
								}
						//}
                    }
                    else
                       time = setTimeout("Repeat()",100); 
                }
            } 
    </script>

    <!-- DEXTUploadX ActiveX 로딩, 아이템 및 프로퍼티 등 설정 -->

 	<script language="javascript">
            function OnLoading1()
			{
				// Post 방식일 경우 아래와 같이 UploadURL 속성에 Post Script 파일을 명시해야 합니다.
				// 표준 포트 외의 다른 포트를 사용하시려면 
				// http://Localhost:8080/DEXTUploadX/Upload/PostScript.asp
				// 와 같이 일반적인 주소 지정 방법과 동일하게 사용하시면 됩니다.

				//alert("start");

				//document.all["WebWrite_dir"].value = opener.document.all["WebWrite_dir"].value;
				//document.all["WebWrite_content"].value = opener.document.all["WebWrite_content"].value;

				//document.all["subject"].value = opener.document.all["subject"].value;
				//document.all["code"].value = opener.document.all["code"].value;
				//document.all["Qstr"].value = opener.document.all["Qstr"].value;
				//document.all["Ktype"].value = opener.document.all["Ktype"].value;


				//alert(document.all["WebWrite_dir"].value);
				//alert(document.all["WebWrite_content"].value);

				//Send_Writesuccess_New
				//document.all["FileUploadMonitor"].UploadURL = "http://www.dohwa.co.kr//SFG/Net_test/FrontEnd/Home_myoffice_SubModule/Devpia/Upload/ProgressUIUpload/ProgressUIUpload.asp";
				document.all["FileUploadMonitor"].UploadURL = "http://sfg.dohwa.co.kr/Net_test/FrontEnd/Home_myoffice_SubModule/Mis/WebWrite_asp/ProgressUIUpload.asp";
				// 파일 매니저의 몇 가지 속성들(DefaultPath, Filter 등등)의 값을 파일 모니터에 복사합니다.
				document.all["FileUploadMonitor"].Properties = opener.document.all["FileUploadManager"].Properties;
				// 이 페이지의 부모 페이지에 있는 파일 매니저 컨트롤의 모든 파일 및 폼 아이템을 파일 모니터 컨트롤에 복사합니다.
				document.all["FileUploadMonitor"].Items = opener.document.all["FileUploadManager"].Items;		
				

				//document.all("FileUploadMonitor")("WebWrite_dir")
				//document.all("FileUploadMonitor")("WebWrite_content")

				//document.all("FileUploadMonitor")("subject")
				//document.all("FileUploadMonitor")("code")
				//document.all("FileUploadMonitor")("Qstr")
				//document.all("FileUploadMonitor")("Ktype")


					if(0 == document.all("FileUploadMonitor").Count) {
							alert(document.all("FileUploadMonitor").EnableEmptyFileUpload);
                           document.all("FileUploadMonitor").EnableEmptyFileUpload = true
                           document.all("FileUploadMonitor").Transfer
						   alert("empty1");
					}				
				
				Pcode = document.all("FileUploadMonitor")("code");
				PQstr = document.all("FileUploadMonitor")("Qstr");
				PKtype = document.all("FileUploadMonitor")("Ktype");


					//alert("WebWrite_content : " + document.all("FileUploadMonitor")("WebWrite_content"));
					//alert("qqq : " + document.all("FileUploadMonitor")("qqq")); 
	        		//alert("WebWrite_dir : " + document.all("FileUploadMonitor")("WebWrite_dir")); 
	        		       
	        		//       alert("subject : " + document.all("FileUploadMonitor")("subject"));


				// DEXTUploadX 전송 정보 초기화
				Init();
				//document.all["FileUploadMonitor"].CheckAutoCloseWindow = true; 


				btnTransfer_Onclick();
			}
    </script>	




	        <SCRIPT LANGUAGE="VBS">
	        	sub OnLoading()
				' Post 방식일 경우 아래와 같이 UploadURL 속성에 Post Script 파일을 명시해야 합니다.
				' 표준 포트 외의 다른 포트를 사용하시려면 
				' http://Localhost:8080/DEXTUploadX/Upload/PostScript.asp
				' 과 같이 일반적인 주소 지정 방법과 동일하게 사용하시면 됩니다.
        			document.all("FileUploadMonitor").UploadURL = "http://sfg.dohwa.co.kr/Net_test/FrontEnd/Home_myoffice_SubModule/ezDoHwaBoard/save.asp"
				' 파일 매니저의 몇 가지 속성들(DefaultPath, Filter 등등)의 값을 파일 모니터에 복사합니다.
		       		document.all("FileUploadMonitor").Properties = opener.document.all("FileUploadManager").Properties
				' 이 페이지의 부모 페이지에 있는 파일 매니저 컨트롤의 모든 파일 및 폼 아이템을 파일 모니터 컨트롤에 복사합니다.
	        		document.all("FileUploadMonitor").Items = opener.document.all("FileUploadManager").Items


				' 실제로 전송 할 파일이 없어서 폼 아이템만 전송해야 한다면 사용자 동의없이 바로 전송한다. 

					'Pcode = document.all("FileUploadMonitor")("code")
					'PQstr = document.all("FileUploadMonitor")("Qstr")
					PKtype = document.all("FileUploadMonitor")("Site_Code")

					If 0 = document.all("FileUploadMonitor").Count then

						MsgBox document.all("FileUploadMonitor").EnableEmptyFileUpload
                           document.all("FileUploadMonitor").EnableEmptyFileUpload = TRUE
                           document.all("FileUploadMonitor").Transfer
						MsgBox document.all("FileUploadMonitor").EnableEmptyFileUpload


					   'window.opener.close()	
					   window.close()
							
							if PKtype = "1" then
								window.opener.opener.location.href = "../../ezDoHwaBoard/Right_JaRyo_Upmubogo.asp?code=" + Pcode + "&Qstr=" + PQstr 
							else								
								window.opener.opener.location.href = "../../ezDoHwaBoard/Right_JaRyo.asp?code=" + Pcode + "&Qstr=" + PQstr 					
							end if								
					else

				Init()
				btnTransfer_Onclick()
                    End If
				
				'MsgBox Pcode
				'MsgBox PQstr
				'MsgBox PKtype
					'MsgBox "Text1 : " & document.all("FileUploadMonitor")("text1")+ _
	        		'       "Text2 : " + document.all("FileUploadMonitor")("text2") + _
	        		'       "UserAddText: " + document.all("FileUploadMonitor")("UserAddText")

	        	end sub
		</SCRIPT>











    <!-- DEXTUploadX 이벤트 코드 -->

    <script language="javascript" for="FileUploadMonitor" event="OnTransferComplete()">          
    </script>

    <script language="javascript" for="FileUploadMonitor" event="OnTransferCancel()"> 
             Repeat(false);
             //alert("전송이 취소되었습니다.");
    </script>

    <!-- DEXTUploadX 버튼 동작 코드 -->

    <script language="javascript">
            
            // 파일 추가 버튼
			function btnAddFile_Onclick()
			{
			    // 전송중이 아닐때만 파일을 추가한다. 
			    if(false == g_bTransfer) { 
			        document.all["FileUploadMonitor"].OpenFileDialog(); 
			        Refresh(); 
			    }
			}
			
			// 항목 삭제 버튼
			function btnDeleteItem_Onclick()
			{
			    // 전송중이 아닐때만 파일을 추가한다. 
			    if(false == g_bTransfer) { 
			        document.all["FileUploadMonitor"].DeleteSelectedFile(); 
			        Refresh(); 
			    }
			}
			


            // 전송 버튼
			function btnTransfer_Onclick()
			{

					if(0 == document.all("FileUploadMonitor").Count) {
						

					}	else {

							// 전송중이 아닐때만 Transfer() 메소드를 호출한다. 
							if(false == g_bTransfer) { 
								document.all["FileUploadMonitor"].Transfer();
								
								if(0 != document.all["FileUploadMonitor"].Count) {
									Repeat(true); 
									
									// 전송상태를 전송중으로 설정하고, 이미지 변경
									g_bTransfer = true; 
									document.images["Transfer"].src="../../img/New_image/btn_ssend_dis.gif"; 
									document.images["Cancel"].src="../../img/New_image/btn_cancel_nor.gif";
									document.images["AddFile"].src="../../img/New_image/btn_fileadd_dis.gif";
									document.images["DeleteItem"].src="../../img/New_image/btn_listdel_dis.gif";
									document.images["MoveUp"].src="../../img/New_image/btn_goup_dis.GIF";
									document.images["MoveDown"].src="../../img/New_image/btn_godown_dis.GIF";
								}
								else {
									// 전송상태를 전송중으로 설정하고, 이미지 변경
									g_bTransfer = true;
									document.images["Transfer"].src="../../img/New_image/btn_ssend_dis.gif";
									document.images["AddFile"].src="../../img/New_image/btn_fileadd_dis.gif";
									document.images["DeleteItem"].src="../../img/New_image/btn_listdel_dis.gif";
									document.images["MoveUp"].src="../../img/New_image/btn_goup_dis.GIF";
									document.images["MoveDown"].src="../../img/New_image/btn_godown_dis.GIF";
								}
							}		      
				
				}
			}
			


			// 취소 버튼
			function btnCancel_Onclick()
			{   
			    document.all["FileUploadMonitor"].Cancel();   
			    Repeat(false);   
			    
			    if(true == g_bTransfer) {
			        // 전송상태를 중지로 선택하고, 이미지 변경
			        g_bTransfer = false; 
			        document.images["Transfer"].src="../../img/New_image/btn_ssend_nor.gif"; 
			        document.images["Cancel"].src="../../img/New_image/btn_close_nor.gif"; 
			        document.images["AddFile"].src="../../img/New_image/btn_fileadd_nor.gif";
	                document.images["DeleteItem"].src="../../img/New_image/btn_listdel_nor.gif";
	                document.images["MoveUp"].src="../../img/New_image/btn_goup_nor.GIF";
	                document.images["MoveDown"].src="../../img/New_image/btn_godown_nor.GIF";
			    }
			}
			
			//파일 위치 이동
			function btnMoveFileUp()
			{ 
			    // 전송중이 아닐때만 파일을 추가한다. 
			    if(false == g_bTransfer) {
			        document.all["FileUploadMonitor"].MoveFileUp(); 
			    }
			}   
					
			function btnMoveFileDown()
			{ 
			    // 전송중이 아닐때만 파일을 추가한다. 
			    if(false == g_bTransfer) {
			        document.all["FileUploadMonitor"].MoveFileDown(); 
			    }
			}
			
			// 체크박스 코드
			function btnCheckbox_Click()
			{
			    if(true == document.all["FileUploadMonitor"].CheckAutoCloseWindow) 
			    {
			        document.all["FileUploadMonitor"].CheckAutoCloseWindow = false; 
			        document.images["Chk_Box"].src="../../img/New_image/btn_checkno.GIF"; 
			    } 
			    else
			    {
			        document.all["FileUploadMonitor"].CheckAutoCloseWindow = true; 
			        document.images["Chk_Box"].src="../../img/New_image/btn_checkyes.GIF"; 
			    }  
			}
    </script>
    
    
<style type="text/css">
<!--
body, td{
	font-family : 돋움;
	font-size : 11px;
}
.style3 {color: #787878;
	font-size : 12px;
}
.style4 {color: #FFFFFF;
	font-size : 12px;
}
-->
</style>

<script type="text/javascript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>

<body onload="OnLoading()" leftmargin="0" topmargin="0" onLoad="MM_preloadImages('../../img/New_image/btn_fileadd_over.gif','../../img/New_image/btn_listdel_over.gif','../../img/New_image/btn_goup_over.GIF','../../img/New_image/btn_godown_over.GIF','../../img/New_image/btn_close_over.gif','../../img/New_image/btn_ssend_over.gif')">



<table width="490" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="../../img/New_image/fum_box01.gif" width="2" height="2"></td>
    <td background="../../img/New_image/fum_box02.gif" width="463" height="2"></td>
    <td><img src="../../img/New_image/fum_box03.gif" width="2" height="2"></td>
  </tr>
  <tr>
    <td background="../../img/New_image/fum_box04.gif" width="2" height="200"></td>
    <td align="center" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="center" background="../../img/New_image/p_head.gif" width="486" height="64" style="padding:4px 15px 0px 17px"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="3"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td style="padding: 0px 0px 2px 0px;"><img src="../../img/New_image/icon_02.GIF" width="4" height="4"></td>
                  <td style="padding: 0px 0px 2px 0px;"><img src="../../img/New_image/text_filename.GIF" width="42" height="10" align="absmiddle"></td>
                  <td colspan="7">
                    <!-- 파일이름 -->
                    <span id="_CurrentFileName" class="style4"></span>
                  </td>
                  </tr>
                <tr>
                  <td colspan="4" height="8"></td>
                </tr>
                <tr>
                  <td width="2%" style="padding: 0px 0px 2px 0px;"><img src="../../img/New_image/icon_02.GIF" width="4" height="4"></td>
                  <td width="12%" style="padding: 0px 0px 2px 0px;"><img src="../../img/New_image/text_timeleft.GIF" width="43" height="10"></td>
                  <td width="24%">
                    <!-- 남은시간 -->
                    <span id="_RemainingTime" class="style4"></span>
                  </td>
                  <td width="2%" style="padding: 0px 0px 2px 0px;"><img src="../../img/New_image/icon_02.GIF" width="4" height="4"></td>
                  <td width="12%" style="padding: 0px 0px 2px 0px;"><img src="../../img/New_image/text_filenum.GIF" width="42" height="10"></td>
                  <td width="16%">
                    <!-- 파일개수 -->
                    <span id="_CurrentCount" class="style4"></span><span class="style4">/</span><span id="_TotalCount" class="style4"></span></td>
                  <td width="2%" style="padding: 0px 0px 2px 0px;"><img src="../../img/New_image/icon_02.GIF" width="4" height="4"></td>
                  <td width="12%" style="padding: 0px 0px 2px 0px;"><img src="../../img/New_image/text_speed.GIF" width="42" height="10"></td>
                  <td width="18%">
                    <!-- 전송속도 -->
                    <span id="_BytesPerSec" class="style4"></span>
                  </td>
                </tr>
              </table>
                <span class="style4"></span></td>
          </tr>
        </table>
        </td>
      </tr>
      <tr>
        <td align="center" bgcolor="#eaecec" style="padding: 10px 10px 10px 10px;"><table width="100%" height="70" border="0" cellpadding="0" cellspacing="1" bgcolor="#c0c5cc">
          <tr>
            <td align="center" bgcolor="#FFFFFF" style="padding: 1px 1px 1px 1px;"><table width="100%" height="70" border="0" cellpadding="0" cellspacing="1" bgcolor="#e9e9e9">
              <tr>
                <td align="center" bgcolor="#FFFFFF" style="padding: 0px 14px 0px 14px;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="2%"><img src="../../img/New_image/icon_03.GIF" width="4" height="4"></td>
                    <td width="14%" valign="top"><img src="../../img/New_image/text_nowfile.GIF" width="55" height="11"></td>
                    <td width="11%" valign="bottom">
                        <!-- 현재파일 -->
                        <span id="_TransferedSizeOfCurrentFile" class="style3"></span>
                    </td>
                    <td width="73%"><table width="100%" height="11" border="0" cellpadding="0" cellspacing="1" bgcolor="#7b8996">
                      <tr>
                        <td bgcolor="#f6f4e0">
                            <!-- Progress 객체를 사용한 현재 파일의 사이즈에 대한 프로그래스바 -->
							    <object id="Progress_file" height="9" width="313"	classid="CLSID:253B8695-AFE6-4918-AE1C-003FCF070D08" codeBase="http://sfg.dohwa.co.kr/DEXTUploadX/DEXTUploadX.cab#version=2,8,2,0" viewastext>
								    <!--프로그래스 바 및 배경색 지정-->
								    <param name="ProgressBarBkColor" value="#f6f4e0" /> 
								    <param name="ProgressBarColor" value="#ff8f0e" />
							    </object>
                        </td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td colspan="4" height="10"></td>
                  </tr>
                  <tr>
                    <td><img src="../../img/New_image/icon_03.GIF" width="4" height="4"></td>
                    <td valign="top"><img src="../../img/New_image/text_allfile.GIF" width="55" height="11"></td>
                    <td valign="bottom" class="style3">
                        <!-- 전체파일 전송량 -->
                        <span id="_TransferedSizeOfTotalFile" class="style3"></span>
                    </td>
                    <td><table width="100%" height="11" border="0" cellpadding="0" cellspacing="1" bgcolor="#7b8996">
                      <tr>
                        <td bgcolor="#f6f4e0">
                            <!-- Progress 객체를 사용한 전체 파일 사이즈에 대한 프로그래스바 -->
							    <object id="Progress_total" height="9" width="313"	classid="CLSID:253B8695-AFE6-4918-AE1C-003FCF070D08"codeBase="http://sfg.dohwa.co.kr/DEXTUploadX/DEXTUploadX.cab#version=2,8,2,0" viewastext>
								    <!--프로그래스 바 및 배경색 지정-->
								    <param name="ProgressBarBkColor" value="#f6f4e0" /> 
								    <param name="ProgressBarColor" value="#ff8f0e" />
							    </object>
                        </td>
                      </tr>
                    </table></td>
                  </tr>
                </table></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td align="center" bgcolor="#eaecec" style="padding: 0px 10px 0px 10px;"><table width="100%" height="110" border="0" cellpadding="0" cellspacing="1" bgcolor="#c0c5cc">
          <tr>
            <td align="center" bgcolor="#FFFFFF">
                <object id="FileUploadMonitor" height="108" width="464" classid="CLSID:96A93E40-E5F8-497A-B029-8D8156DE09C5"
					codeBase="http://sfg.dohwa.co.kr/DEXTUploadX/DEXTUploadX.cab#version=2,8,2,0" viewastext>
					<param name="DialogBoxMode" value="DLG_LISTVIEW" />
				</object>
            </td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="32" align="center" bgcolor="#eaecec" style="padding: 4px 10px 0px 10px;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td valign="top">
                <!-- 파일 추가 버튼 -->
                <!--img src="../../img/New_image/btn_fileadd_nor.gif" name="AddFile" width="80" height="23" border="0" onclick="btnAddFile_Onclick()"-->
                <!-- 항목 삭제 버튼 -->
                <!--img src="../../img/New_image/btn_listdel_nor.gif" name="DeleteItem" width="80" height="23" border="0" onclick="btnDeleteItem_Onclick()"-->
            </td>
            <td align="right" valign="top">
                <!-- 파일 한칸 위로 이동 -->
                <!--img src="../../img/New_image/btn_goup_nor.GIF" name="MoveUp" width="21" height="20" border="0" onclick="btnMoveFileUp()"-->
                <!--파일 한칸 아래로 이동-->
                <!--img src="../../img/New_image/btn_godown_nor.GIF" name="MoveDown" width="21" height="20" border="0" onclick="btnMoveFileDown()"-->
            </td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td height="1" align="center" bgcolor="#a7b2ba"></td>
      </tr>
      <tr>
        <td height="32" align="center" bgcolor="#969fa2" style="padding: 5px 10px 0px 10px;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="4%">
                <!-- 체크박스 -->
                <!--img src="../../img/New_image/btn_checkno.GIF" name="Chk_Box" width="14" height="14" onclick="btnCheckbox_Click()"-->
            </td>
            <td width="28%"><!--img src="../../img/New_image/text_closeaftersend.gif" width="90" height="10"--></td>
            <td width="68%" align="right">
                <!-- 전송버튼 -->
                <!--img src="../../img/New_image/btn_ssend_nor.gif" name="Transfer" width="80" height="23" border="0" onclick="btnTransfer_Onclick()"-->
                <!-- 닫기 버튼 -->
                <!--img src="../../img/New_image/btn_close_nor.gif" name="Cancel" width="80" height="23" border="0" onclick="btnCancel_Onclick()"-->
            </td>
          </tr>
        </table></td>
      </tr>
    </table></td>
    <td background="../../img/New_image/fum_box05.gif" width="2" height="200">저장 중에 창을 닫으시면 저장이 안됩니다.</td>
  </tr>
  <tr>
    <td><img src="../../img/New_image/fum_box10.gif" width="2" height="2"></td>
    <td background="../../img/New_image/fum_box09.gif" height="2"></td>
    <td><img src="../../img/New_image/fum_box11.gif" width="2" height="2"></td>
  </tr>
</table>




				<!--input type="text" NAME="WebWrite_dir">
				<div style="display:none;" name="dtxt" id="dtxt">
					<textarea name="WebWrite_content"></textarea>
				</div>

				<input type="text" name="subject">
				<input type="text" name="code">
				<input type="text" name="Qstr">
				<input type="text" name="Ktype"-->	

 

 </body>
</html>
