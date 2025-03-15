<%@ LANGUAGE="VBSCRIPT"	CODEPAGE="949"%>
<!--#include file="../../../../dbopen.asp"-->
<%
	comboList =""
	
	IF request("I_GUBUN") = "1" Then 
		sql = " SELECT DISTINCT SAP_GROUP, Group_NM FROM interface_gu_name_code ORDER BY Group_NM "
		Set rs= Server.CreateObject("ADODB.Recordset")
		rs.CursorType=1 ' CursorTyp : 0 : adOpenForwardOnly (기본값) 	1 : adOpenKeyset 	2 : adOpenDynamic 	3 : adOpenStatic 
		rs.Open sql, DbCon	
		
		'response.write "rs.Recordcount=" & rs.Recordcount & "<br>"
		
		If rs.Recordcount <> 0 Then 
			For i = 1 to rs.Recordcount
				SAP_GROUP = rs("SAP_GROUP")
				Group_NM = rs("Group_NM")
				if Group_NM<>"" Then 
					If i = 1 Then
						'comboList = comboList  & SAP_GROUP & "|" & Group_NM & "^|^"
						comboList =  "^|^,"  & SAP_GROUP & "|" & Group_NM & "^|^"
					ElseIf i = rs.Recordcount Then
						comboList = comboList & ", " & SAP_GROUP & "|" & Group_NM 
					Else 
						comboList = comboList & ", " & SAP_GROUP & "|" & Group_NM & "^|^"
					End if	
				End If
				rs.MoveNext			
			Next
		End If
		rs.Close
		Set rs = Nothing	
	ElseIF request("I_GUBUN") = "2" Then 
		sql = " SELECT SAP_GROUP, SAP_GROUP_SUB, Sub_NM  FROM interface_gu_name_code where SAP_GROUP='" & request("SAP_GROUP") &"' "
		Set rs= Server.CreateObject("ADODB.Recordset")
		rs.CursorType=1 ' CursorTyp : 0 : adOpenForwardOnly (기본값) 	1 : adOpenKeyset 	2 : adOpenDynamic 	3 : adOpenStatic 
		rs.Open sql, DbCon	
		
		'response.write "rs.Recordcount=" & rs.Recordcount & "<br>"
		
		If rs.Recordcount <> 0 Then 
			For i = 1 to rs.Recordcount
				SAP_GROUP_SUB = rs("SAP_GROUP_SUB")
				Sub_NM = rs("Sub_NM")
				if Sub_NM<>"" Then 
					If i = 1 Then
						'comboList = comboList  & SAP_GROUP_SUB & "|" & Sub_NM & "^|^"
						comboList = "^|^,"  & SAP_GROUP_SUB & "|" & Sub_NM & "^|^"
					ElseIf i = rs.Recordcount Then
						comboList = comboList & ", " & SAP_GROUP_SUB & "|" & Sub_NM 
					Else 
						comboList = comboList & ", " & SAP_GROUP_SUB & "|" & Sub_NM & "^|^"						
					End if	
				End If	
				rs.MoveNext			
			Next
		End If
		rs.Close
		Set rs = Nothing	
	End IF
	
	response.Charset = "euc-kr"
	response.write comboList & vbCrLF
%>

