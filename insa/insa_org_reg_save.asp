<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
On Error Resume Next
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
Dim u_type, reg_user, mod_user
Dim org_level, org_code, org_date, org_empno
Dim org_company, org_bonbu, org_saupbu, org_team, org_empname
Dim org_cost_group, org_cost_center, org_reside_company
Dim owner_org, owner_orgname, owner_empno, owner_empname
Dim org_table_org, org_zip, org_sido, org_gugun, org_dong
Dim org_addr, org_end_date, tel_ddd, tel_no1, tel_no2
Dim org_reside_place, org_reg_date, org_mod_date
Dim rsEmpCnt, trade_code
Dim rs_stock, stock_end_date, stock_level, end_msg

u_type = Request.Form("u_type")

org_level = Request.Form("org_level")
org_code = Request.Form("org_code")
org_name = Request.Form("org_name")
org_date = f_toString(Request.Form("org_date"), "0000-00-00")
org_empno = Request.Form("org_empno")
org_empname = Request.Form("org_empname")
org_company = Request.Form("org_company")
org_bonbu = f_toString(Request.Form("org_bonbu"), "")
org_saupbu = f_toString(Request.Form("org_saupbu"), "")
org_team = f_toString(Request.Form("org_team"), "")

org_cost_group = f_toString(Request.Form("org_cost_group"), "")

org_cost_center = Request.Form("org_cost_center")

org_reside_company = f_toString(Request.Form("org_reside_company"), "")

owner_org = Request.Form("owner_org")
owner_orgname = Request.Form("owner_orgname")
owner_empno = Request.Form("owner_empno")
owner_empname = Request.Form("owner_empname")
org_table_org = Int(Request.Form("org_table_org"))
org_zip = Request.Form("org_zip")
org_sido = Request.Form("org_sido")
org_gugun = Request.Form("org_gugun")
org_dong = Request.Form("org_dong")
org_addr = Request.Form("org_addr")
org_end_date = f_toString(Request.Form("org_end_date"), "0000-00-00")
tel_ddd = Request.Form("tel_ddd")
tel_no1 = Request.Form("tel_no1")
tel_no2 = Request.Form("tel_no2")

org_reside_place = Request.Form("org_reside_place")

org_reg_date = Request.Form("org_reg_date")
org_mod_date = Request.Form("org_mod_date")

'거래처 코드 추가
trade_code = Request.Form("trade_code")

Select Case org_level
	Case "회사" : org_company = org_name
	Case "본부" : org_bonbu = org_name
	Case "사업부" : org_saupbu = org_name
	Case "팀" : org_team = org_name
End Select

If tel_ddd = "" Then
   tel_ddd = ""
   tel_no1 = ""
   tel_no2 = ""
End If

If org_level = "상주처" Then
	org_cost_center = "상주직접비"
Else
'	org_cost_group = org_saupbu
	org_cost_group = org_bonbu

	If org_saupbu = "" Then
		If org_bonbu = ""  Then
			org_cost_group = org_company
		Else
			org_cost_group = org_bonbu
		End If
	End If
End If

'???
'If org_reside_company <> "" Then
'   org_cost_group = request.form("org_cost_group")
'end if

DBConn.BeginTrans

If u_type = "U" Then
	'조직 정보 변경
	objBuilder.Append "UPDATE emp_org_mst SET "
	objBuilder.Append "	org_level='"&org_level&"',org_company='"&org_company&"',org_bonbu='"&org_bonbu&"', "
	objBuilder.Append "	org_saupbu='"&org_saupbu&"',org_team='"&org_team&"',org_name='"&org_name&"', "
	objBuilder.Append "	org_reside_place='"&org_reside_place&"',org_reside_company='"&org_reside_company&"', "
	objBuilder.Append "	org_cost_group='"&org_cost_group&"', "
	objBuilder.Append "	org_empno='"&org_empno&"',org_emp_name='"&org_empname&"',org_date='"&org_date&"', "
	objBuilder.Append "	org_tel_ddd='"&tel_ddd&"',org_tel_no1='"&tel_no1&"',org_tel_no2='"&tel_no2&"', "
	objBuilder.Append "	org_owner_org='"&owner_org&"',org_owner_empno='"&owner_empno&"',org_owner_empname='"&owner_empname&"', "
	objBuilder.Append "	org_table_org='"&org_table_org&"',org_sido='"&org_sido&"',org_gugun='"&org_gugun&"', "
	objBuilder.Append "	org_dong='"&org_dong&"',org_addr='"&org_addr&"',org_cost_group='"&org_cost_group&"', "
	objBuilder.Append "	org_cost_center='"&org_cost_center&"',org_end_date='"&org_end_date&"',org_mod_date=NOW(), "
	objBuilder.Append "	org_mod_user='"&user_name&"', trade_code = '"&trade_code&"' "
	objBuilder.Append "WHERE org_code = '"&org_code&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'조직 코드가 인사 마스터에 사용된 경우에만 인사 마스터 정보 일괄 수정[허정호_20210730]
	objBuilder.Append "SELECT COUNT(*) AS emp_cnt FROM emp_master "
	objBuilder.Append "WHERE emp_pay_id <> '2' "
	objBuilder.Append "	AND emp_org_code = '"&org_code&"' "

	Set rsEmpCnt = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If CInt(rsEmpCnt("emp_cnt")) > 0 Then
		'인사 마스터 정보 일괄 변경
		objBuilder.Append "UPDATE emp_master SET "
		objBuilder.Append "	emp_company = '"&org_company&"',"
		objBuilder.Append "	emp_bonbu = '"&org_bonbu&"',"
		objBuilder.Append "	emp_saupbu = '"&org_saupbu&"',"
		objBuilder.Append "	emp_team = '"&org_team&"',"
		objBuilder.Append "	emp_org_name = '"&org_name&"',"
		objBuilder.Append "	emp_reside_company = '"&org_reside_company&"',"
		objBuilder.Append "	emp_reside_place = '"&org_reside_place&"',"
		objBuilder.Append "	cost_center = '"&org_cost_center&"',"
		objBuilder.Append "	cost_group = '"&org_cost_group&"',"
		objBuilder.Append "	emp_mod_user = '"&user_name&"',"
		objBuilder.Append "	emp_mod_date = NOW() "
		objBuilder.Append "WHERE emp_org_code = '"&org_code&"' "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		'회원 정보 일괄 변경
		objBuilder.Append "UPDATE memb SET "
		objBuilder.Append "	emp_company = '"&org_company&"',"
		objBuilder.Append "	bonbu = '"&org_bonbu&"',"
		objBuilder.Append "	saupbu = '"&org_saupbu&"',"
		objBuilder.Append "	team = '"&org_team&"',"
		objBuilder.Append "	org_name = '"&org_name&"',"
		objBuilder.Append "	reside_company = '"&org_reside_company&"',"
		objBuilder.Append "	reside_place = '"&org_reside_place&"',"
		objBuilder.Append "	reg_name = '"&user_name&"',"
		objBuilder.Append "	mod_date = NOW() "
		objBuilder.Append "WHERE user_id IN (SELECT emp_no FROM emp_master WHERE emp_org_code = '"&org_code&"') "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	End If

	rsEmpCnt.Close() : Set rsEmpCnt = Nothing

Else
	objBuilder.Append "INSERT INTO emp_org_mst(org_code,org_level,org_company,org_bonbu,org_saupbu, "
	objBuilder.Append "org_team,org_name,org_reside_place,org_reside_company,org_cost_group, "
	objBuilder.Append "org_empno,org_emp_name,org_date,org_tel_ddd,org_tel_no1, "
	objBuilder.Append "org_tel_no2,org_cost_center,org_owner_org,org_owner_empno,org_owner_empname, "
	objBuilder.Append "org_table_org,org_sido,org_gugun,org_dong,org_addr, "
	objBuilder.Append "org_reg_date,org_reg_user, trade_code)values("
	objBuilder.Append "'"&org_code&"','"&org_level&"','"&org_company&"','"&org_bonbu&"','"&org_saupbu&"', "
	objBuilder.Append "'"&org_team&"','"&org_name&"','"&org_reside_place&"','"&org_reside_company&"','"&org_cost_group&"', "
	objBuilder.Append "'"&org_empno&"','"&org_empname&"','"&org_date&"','"&tel_ddd&"','"&tel_no1&"', "
	objBuilder.Append "'"&tel_no2&"','"&org_cost_center&"','"&owner_org&"','"&owner_empno&"','"&owner_empname&"', "
	objBuilder.Append "'"&org_table_org&"','"&org_sido&"','"&org_gugun&"','"&org_dong&"','"&org_addr&"', "
	objBuilder.Append "NOW(),'"&user_name&"', '"&trade_code&"')"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

' 창고코드 등록
If org_level = "본사" Or org_level = "팀" Then
	If org_code <> "" Or org_code <> " " Then
		objBuilder.Append "SELECT stock_level "
		objBuilder.Append "FROM met_stock_code "
		objBuilder.Append "WHERE stock_code = '"&org_code&"' "

		Set rs_stock = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_stock.EOF Then
			stock_end_date = "1900-01-01"

			If org_level = "회사" Then
				stock_level = "본사"
			Else
				stock_level = "팀"
			End If

			objBuilder.Append "INSERT INTO met_stock_code("
			objBuilder.Append "	stock_code,stock_level,stock_name,stock_company,stock_bonbu, "
			objBuilder.Append "	stock_saupbu,stock_team,stock_open_date,stock_end_date,stock_manager_code, "
			objBuilder.Append "	stock_manager_name, reg_date,reg_user)VALUES("
			objBuilder.Append "'"&org_code&"','"&stock_level&"','"&org_name&"','"&org_company&"','"&org_bonbu&"', "
			objBuilder.Append "'"&org_saupbu&"','"&org_team&"','"&org_date&"','"&stock_end_date&"','"&org_empno&"', "
			objBuilder.Append "'"&org_empname&"',now(),'"&reg_user&"') "
		Else
			objBuilder.Append "UPDATE met_stock_code SET "
			objBuilder.Append " stock_name = '"&org_name&"', "
			objBuilder.Append " stock_company = '"&org_company&"', "
			objBuilder.Append " stock_bonbu = '"&org_bonbu&"', "
			objBuilder.Append " stock_saupbu = '"&org_saupbu&"', "
			objBuilder.Append " stock_team = '"&org_team&"', "
			objBuilder.Append " stock_open_date = '"&org_date&"', "
			objBuilder.Append " stock_manager_code = '"&org_empno&"', "
			objBuilder.Append " stock_manager_name = '"&org_empname&"' "
			objBuilder.Append "WHERE stock_code='"&org_code&"' "
		End If

		rs_stock.Close() : Set rs_stock = Nothing

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	End If
End If

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "등록 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans

	If u_type = "U" Then
		end_msg = "정상적으로 수정되었습니다."
	Else
		end_msg = "정상적으로 등록되었습니다."
	End If
End If

DBConn.Close() : Set DBConn = Nothing

Response.write "<script type='text/javascript'>"
Response.write "	alert('"&end_msg&"');"
'Response.write "	location.replace('insa_org.asp');"
Response.write "	self.opener.location.reload();"
Response.write "	window.close();"
Response.write "</script>"
Response.End
%>
