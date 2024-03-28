<%@LANGUAGE="VBSCRIPT"%>
<%Response.ContentType="text/html;charset=euc-kr"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim result : result = ""
Dim sql
Dim orgColumn, sqlWhere
Dim srchType, srchCompany, srchBonbu, srchSaupbu, srchTeam

srchType			= replaceXSS(request("srchType"))		'//검색구분(1:회사,2:본부,3:사업부,4:팀,5:상주처)
srchCompany	= replaceXSS(request("srchCompany"))	'//회사명
srchBonbu		= replaceXSS(request("srchBonbu"))		'//본부명
srchSaupbu		= replaceXSS(request("srchSaupbu"))	'//사업부명
srchTeam			= replaceXSS(request("srchTeam"))		'//팀명

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

If srchType="" Or srchType="1" Then
	orgColumn = "org_company"
	sqlWhere = " AND org_level = '회사' "
ElseIf srchType="" Or srchType="2" Then
	orgColumn = "org_bonbu"
	sqlWhere = " AND org_level = '본부' AND org_company='" & srchCompany & "' "
ElseIf srchType="" Or srchType="3" Then
	orgColumn = "org_saupbu"
	sqlWhere = " AND org_level = '사업부' AND org_company='" & srchCompany & "' AND org_bonbu='" & srchBonbu & "' "
ElseIf srchType="" Or srchType="4" Then
	orgColumn = "org_team"
	sqlWhere = " AND org_level = '팀' AND org_company='" & srchCompany & "' AND org_bonbu='" & srchBonbu & "' AND org_saupbu='" & srchSaupbu & "' "
ElseIf srchType="" Or srchType="5" Then
	orgColumn = "org_name"
	sqlWhere = " AND org_level = '상주처' AND org_company='" & srchCompany & "' AND org_bonbu='" & srchBonbu & "' AND org_saupbu='" & srchSaupbu & "' AND org_team='" & srchTeam & "' "
End If

'//get company list
sql="SELECT org_code, " & orgColumn & " FROM emp_org_mst WHERE isNull(org_end_date) " & sqlWhere & " ORDER BY org_code ASC"

Rs.Open Sql, Dbconn, 1

If Not(Rs.BOF Or Rs.EOF) Then
	Do Until Rs.EOF
		If Trim(result&"")<>"" Then result = result & ","
		result = result & "{""orgCode"":""" & Rs("org_code") & """"
		result = result & ",""orgCompany"":""" & Rs("org_company") & """}"
		Rs.movenext
	Loop
End If

Rs.close : Set Rs = Nothing
Dbconn.close : Set Dbconn = Nothing

'//return json data
If Trim(result&"")<>"" Then
	result = "{""result"":[" & result & "]}"
End If

Response.write result

%>