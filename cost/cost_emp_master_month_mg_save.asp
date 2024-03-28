<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
Response.expires=-1
Response.ContentType = "application/json"
Response.Charset = "euc-kr"

'On Error Resume Next

'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim result : result = "fail"

Dim empMonth, empNo, costCenter, empOrgCode, empOrgName
Dim empCompany, empBonbu, empSaupbu, empTeam, costGroup
Dim costOrgCode, costOrgName, rsOrgName

empMonth	= replaceXSS(Unescape(toString(request("empMonth"),"")))
empNo		= replaceXSS(Unescape(toString(request("empNo"),"")))
costCenter	= replaceXSS(Unescape(toString(request("costCenter"),"")))
costGroup	= replaceXSS(Unescape(toString(request("costGroup"),"")))
empOrgCode	= replaceXSS(Unescape(toString(request("empOrgCode"),"")))
empOrgName	= replaceXSS(Unescape(toString(request("empOrgName"),"")))

empCompany	= replaceXSS(Unescape(toString(request("empCompany"),"")))
empBonbu	= replaceXSS(Unescape(toString(request("empBonbu"),"")))
empSaupbu	= replaceXSS(Unescape(toString(request("empSaupbu"),"")))
empTeam		= replaceXSS(Unescape(toString(request("empTeam"),"")))

costOrgCode	= replaceXSS(Unescape(toString(request("costOrgCode"),"")))


'objBuilder.Append "SELECT org_bonbu FROM emp_org_mst WHERE org_code = '"&costOrgCode&"' "
'Set rsOrgName = DBConn.Execute(objBuilder.ToString())
'objBuilder.Clear()

'costOrgName = rsOrgName("org_bonbu")

'rsOrgName.Close() : Set rsOrgName = Nothing

'//파라미터 체크
If empMonth="" Or empNo="" Or costCenter="" Or costGroup="" Then
	result = "invalid"
Else
	'//set 월별 인사
	'sql = "UPDATE emp_master_month SET "
	'sql = sql & "  cost_group='"   & costGroup  & "' "
	'sql = sql & ", cost_center='"  & costCenter & "' "
	'sql = sql & ", emp_company='"  & empCompany & "' "
	'sql = sql & ", emp_bonbu='"    & empBonbu   & "' "
	'sql = sql & ", emp_saupbu='"   & empSaupbu  & "' "
	'sql = sql & ", emp_team='"     & empTeam    & "' "
	'sql = sql & ", emp_org_code='" & empOrgCode & "' "
	'sql = sql & ", emp_org_name='" & empOrgName & "' "
	'sql = sql & " WHERE emp_no='" & empNo & "' AND emp_month='" & empMonth & "' "

	objBuilder.Append "UPDATE emp_master_month SET "
	objBuilder.Append "	cost_group='"   & costGroup  & "', "
	objBuilder.Append "	cost_center='"  & costCenter & "', "
	objBuilder.Append "	emp_company='"  & empCompany & "', "
	objBuilder.Append "	emp_bonbu='"    & empBonbu   & "', "
	objBuilder.Append "	emp_saupbu='"    & empSaupbu   & "', "
	objBuilder.Append "	emp_team='"     & empTeam    & "', "
	objBuilder.Append "	emp_org_code='" & empOrgCode & "', "
	objBuilder.Append "	emp_org_name='" & empOrgName & "' "
	objBuilder.Append "WHERE emp_no='" & empNo & "' AND emp_month='" & empMonth & "' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'//set 월별 급여
	'sql = "UPDATE pay_month_give SET "
	'sql = sql & " cost_group='" & costGroup & "' "
	'sql = sql & ", cost_center='" & costCenter & "' "
	'sql = sql & ", pmg_company='" & empCompany& "' "
	'sql = sql & ", pmg_bonbu='" & empBonbu& "' "
	'sql = sql & ", pmg_saupbu='" & empSaupbu & "' "
	'sql = sql & ", pmg_team='" & empTeam & "' "
	'sql = sql & ", pmg_org_name='" & empOrgName & "' "
	'sql = sql & ", mg_saupbu='" & costGroup & "' "
	'sql = sql & " WHERE pmg_emp_no='" & empNo & "' AND pmg_yymm='" & empMonth & "' "

	objBuilder.Append "UPDATE pay_month_give SET "
	objBuilder.Append "	cost_group='"&costGroup&"', "
	objBuilder.Append "	cost_center='"&costCenter& "', "
	objBuilder.Append "	pmg_company='"&empCompany& "', "
	objBuilder.Append "	pmg_bonbu='"&empBonbu&"', "
	objBuilder.Append "	pmg_saupbu='"&empSaupbu&"', "
	objBuilder.Append "	pmg_team='"&empTeam&"', "
	objBuilder.Append "	pmg_org_name='"&empOrgName&"', "
	objBuilder.Append "	mg_saupbu='"&costGroup&"' "
	objBuilder.Append "WHERE pmg_emp_no='"&empNo&"' AND pmg_yymm='"&empMonth&"' AND pmg_company = '"&empCompany&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	result = "succ"

	DBConn.Close : Set DBConn = Nothing

End If

If Err.number <> 0 Then
	result = "error"
End IF

'//return json data
If Trim(result&"")<>"" Then
	result = "{""result"":""" & result & """}"
End If

Response.write result
%>