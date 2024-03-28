<%@LANGUAGE="VBSCRIPT"%>
<%
Response.expires=-1
Response.ContentType = "application/json"
Response.Charset = "euc-kr"
%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
On Error Resume Next

Dim result : result = "fail"
Dim sql
Dim orgColumn, sqlWhere
Dim empNo, costCenter, costGroup

empMonth	= replaceXSS(Unescape(toString(request("empMonth"),"")))	'//
empNo		= replaceXSS(Unescape(toString(request("empNo"),"")))		'//
costCenter	= replaceXSS(Unescape(toString(request("costCenter"),"")))	'//
costGroup	= replaceXSS(Unescape(toString(request("costGroup"),"")))	'//
empOrgCode	= replaceXSS(Unescape(toString(request("empOrgCode"),"")))	'//
empOrgName	= replaceXSS(Unescape(toString(request("empOrgName"),"")))	'//

empCompany	= replaceXSS(Unescape(toString(request("empCompany"),"")))	'//
empBonbu	= replaceXSS(Unescape(toString(request("empBonbu"),"")))	'//
empSaupbu	= replaceXSS(Unescape(toString(request("empSaupbu"),"")))	'//
empTeam		= replaceXSS(Unescape(toString(request("empTeam"),"")))		'//

'//파라미터 체크
If empMonth="" Or empNo="" Or costCenter="" Or costGroup="" Then
	result = "invalid"
Else

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect

	'//set 월별 인사
	sql = "UPDATE emp_master_month SET "
	sql = sql & "  cost_group='"   & costGroup  & "' "
	sql = sql & ", cost_center='"  & costCenter & "' "
	sql = sql & ", emp_company='"  & empCompany & "' "
	sql = sql & ", emp_bonbu='"    & empBonbu   & "' "
	sql = sql & ", emp_saupbu='"   & empSaupbu  & "' "
	sql = sql & ", emp_team='"     & empTeam    & "' "
	sql = sql & ", emp_org_code='" & empOrgCode & "' "
	sql = sql & ", emp_org_name='" & empOrgName & "' "
	sql = sql & " WHERE emp_no='" & empNo & "' AND emp_month='" & empMonth & "' "
	Dbconn.execute sql

	'//set  인사 ' 2019-04-13 박정신 부장님 요구로 월별 인사를 변경하면 원본도 같이 바뀌도록 수정
	'sql = "UPDATE emp_master SET "
	'sql = sql & "  cost_group='"   & costGroup  & "' "
	'sql = sql & ", cost_center='"  & costCenter & "' "
	'sql = sql & ", emp_company='"  & empCompany & "' "
	'sql = sql & ", emp_bonbu='"    & empBonbu   & "' "
	'sql = sql & ", emp_saupbu='"   & empSaupbu  & "' "
	'sql = sql & ", emp_team='"     & empTeam    & "' "
	'sql = sql & ", emp_org_code='" & empOrgCode & "' "
	'sql = sql & ", emp_org_name='" & empOrgName & "' "
	'sql = sql & " WHERE emp_no='" & empNo & "' "
    'Dbconn.execute sql


	'//set 월별 급여
	sql = "UPDATE pay_month_give SET "
	sql = sql & " cost_group='" & costGroup & "' "
	sql = sql & ", cost_center='" & costCenter & "' "

	sql = sql & ", pmg_company='" & empCompany& "' "
	sql = sql & ", pmg_bonbu='" & empBonbu& "' "
	sql = sql & ", pmg_saupbu='" & empSaupbu & "' "
	sql = sql & ", pmg_team='" & empTeam & "' "
	sql = sql & ", pmg_org_name='" & empOrgName & "' "
	sql = sql & ", mg_saupbu='" & costGroup & "' "
	sql = sql & " WHERE pmg_emp_no='" & empNo & "' AND pmg_yymm='" & empMonth & "' "
	Dbconn.execute sql

	'sql = "UPDATE card_slip SET cost_center='" & costCenter & "', mg_saupbu='" & costGroup & "' WHERE emp_no='" & empNo & "' AND DATE_FORMAT(slip_date,'%Y%m')='" & empMonth & "' "
	'sql = "UPDATE card_slip SET cost_center='" & costCenter & "' WHERE emp_no='" & empNo & "' AND DATE_FORMAT(slip_date,'%Y%m')='" & empMonth & "' "
	'Dbconn.execute sql

	result = "succ"

	Dbconn.close : Set Dbconn = Nothing

End If

If Err.number<>0 Then
	result = "error"
End IF


'//return json data
If Trim(result&"")<>"" Then
	result = "{""result"":""" & result & """}"
End If
Response.write result

%>