<!--#include virtual="/common/inc_top.asp"-->
<%
Response.expires=-1
Response.ContentType = "application/json"
Response.Charset = "euc-kr"
%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
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
'On Error Resume Next

Dim result : result = "fail"
Dim empNo, costCenter, costGroup, empOrgCode, empOrgName
Dim empCompany, empBonbu, empSaupbu, empTeam

Dim orgColumn, sqlWhere

'empMonth	= replaceXSS(Unescape(toString(request("empMonth"),"")))
empNo		= replaceXSS(Unescape(toString(Request("empNo"),"")))
costCenter	= replaceXSS(Unescape(toString(Request("costCenter"),"")))
costGroup	= replaceXSS(Unescape(toString(Request("costGroup"),"")))
empOrgCode	= replaceXSS(Unescape(toString(Request("empOrgCode"),"")))
empOrgName	= replaceXSS(Unescape(toString(Request("empOrgName"),"")))

empCompany	= replaceXSS(Unescape(toString(Request("empCompany"),"")))
empBonbu	= replaceXSS(Unescape(toString(Request("empBonbu"),"")))
empSaupbu	= replaceXSS(Unescape(toString(Request("empSaupbu"),"")))
empTeam		= replaceXSS(Unescape(toString(Request("empTeam"),"")))

'//파라미터 체크
If empNo = "" Or costCenter = "" Or costGroup = "" Then
	result = "invalid"
Else
	objBuilder.Append "UPDATE emp_master SET "
	objBuilder.Append "  cost_group='"   & costGroup  & "' "
	objBuilder.Append ", cost_center='"  & costCenter & "' "
	objBuilder.Append ", emp_company='"  & empCompany & "' "
	objBuilder.Append ", emp_bonbu='"    & empBonbu   & "' "
	objBuilder.Append ", emp_saupbu='"   & empSaupbu  & "' "
	objBuilder.Append ", emp_team='"     & empTeam    & "' "
	objBuilder.Append ", emp_org_code='" & empOrgCode & "' "
	objBuilder.Append ", emp_org_name='" & empOrgName & "' "
	objBuilder.Append " WHERE emp_no='" & empNo & "' "

    DBConn.Execute objBuilder.ToString()
	objBuilder.Clear()

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

	objBuilder.Append "UPDATE memb SET "
	objBuilder.Append "	emp_company = '"&empCompany&"', "
	objBuilder.Append "	bonbu = '"&empBonbu&"', "
	objBuilder.Append "	team = '"&empTeam&"', "
	objBuilder.Append "	org_name = '"&empOrgName&"', "
	objBuilder.Append "	reside_place = '"&&"', "
	objBuilder.Append "	reside_company = '"&&"' "
	objBuilder.Append "WHERE emp_no = '"&empNo&"' "

	DBConn.Execute objBuilder.ToString()
	objBuilder.Clear()

	'sql = "UPDATE card_slip SET cost_center='" & costCenter & "', mg_saupbu='" & costGroup & "' WHERE emp_no='" & empNo & "' AND DATE_FORMAT(slip_date,'%Y%m')='" & empMonth & "' "
	'sql = "UPDATE card_slip SET cost_center='" & costCenter & "' WHERE emp_no='" & empNo & "' AND DATE_FORMAT(slip_date,'%Y%m')='" & empMonth & "' "
	'Dbconn.execute sql

	result = "succ"

	Dbconn.close : Set Dbconn = Nothing

End If

'If Err.number<>0 Then
'	result = "error"
'End IF


'//return json data
If Trim(result&"")<>"" Then
	result = "{""result"":""" & result & """}"
End If
Response.write result

%>