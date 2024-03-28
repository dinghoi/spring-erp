<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
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
Dim u_type, car_no, car_old_no, car_name, car_year, oil_kind
Dim car_owner, buy_gubun, rental_company, car_company, car_reg_date
Dim car_use, car_use_dept, owner_emp_no, owner_emp_name
Dim emp_grade, start_date, car_status, car_comment, last_km
Dim last_check_date, end_date, insurance_company, insurance_date
Dim insurance_amt, emp_user, emp_org_code, emp_org_name

'on Error resume next

u_type = Request.Form("u_type")
car_no = Request.Form("car_no")
car_old_no = Request.Form("car_old_no")
car_name = Request.Form("car_name")
oil_kind = Request.Form("oil_kind")
car_owner = Request.Form("car_owner")
buy_gubun = Request.Form("buy_gubun")
rental_company = Request.Form("rental_company")
car_company = Request.Form("car_company")
car_use = Request.Form("car_use")
car_use_dept = Request.Form("car_use_dept")
owner_emp_no = Request.Form("owner_emp_no")
owner_emp_name = Request.Form("emp_name")
emp_grade = Request.Form("emp_grade")
car_status = Request.Form("car_status")
car_comment = Request.Form("car_comment")
last_km = Int(Request.Form("last_km"))

car_reg_date = Request.Form("car_reg_date")
If car_reg_date = "" Or isnull(car_reg_date) Then
   car_reg_date = "1900-01-01"
End If

start_date = Request.Form("start_date")
If start_date = "" Or IsNull(start_date) Then
   start_date = "1900-01-01"
End If

last_check_date = Request.Form("last_check_date")
If last_check_date = "" Or IsNull(last_check_date) Then
   last_check_date = "1900-01-01"
End If

end_date = Request.Form("end_date")
If end_date = "" Or IsNull(end_date) Then
   end_date = "1900-01-01"
End If

car_year = Request.Form("car_year")
If car_year = "" Or IsNull(car_year) Then
   car_year = "1900-01-01"
End If

'emp_company = Request.Form("emp_company")	'-> 저장된 쿠키 값으로 대체 사용[허정호_20210722]
emp_org_code = Request.Form("emp_org_code")
emp_org_name = Request.Form("emp_org_name")

insurance_company = ""
insurance_date = ""
insurance_amt = 0

DBConn.BeginTrans

'emp_user = Request.Cookies("nkpmg_user")("coo_user_name") -> user_name 쿠키값으로 대체[허정호_20210722]

If u_type = "U" Then
	objBuilder.Append "UPDATE car_info SET "
	objBuilder.Append "	car_name = '"&car_name&"', car_year = '"&car_year&"', oil_kind = '"&oil_kind&"', insurance_amt ='0', "
	objBuilder.Append "	car_owner = '"&car_owner&"', buy_gubun = '"&buy_gubun&"', rental_company = '"&rental_company&"', "
	objBuilder.Append "	car_reg_date = '"&car_reg_date&"', car_use_dept = '"&car_use_dept&"', "
	objBuilder.Append "	car_company = '"&car_company&"', car_use = '"&car_use&"', owner_emp_no = '"&owner_emp_no&"', "
	objBuilder.Append "	owner_emp_name='"&owner_emp_name&"', start_date = '"&start_date&"', end_date = '"&end_date&"', "
	objBuilder.Append "	last_km = '"&last_km&"', last_check_date = '"&last_check_date&"', car_status = '"&car_status&"', "
	objBuilder.Append "	car_comment = '"&car_comment&"', mod_emp_no = '"&user_id&"', mod_emp_name = '"&user_name&"', mod_date = NOW() "
	objBuilder.Append "WHERE car_no = '"&car_no&"'; "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
Else
	'//기등록 여부 체크
	Dim nCarCnt : nCarCnt = 0
	Dim rs_car, end_msg

	objBuilder.Append "SELECT COUNT(*) AS cnt FROM car_info "
	objBuilder.Append "WHERE car_no = '"&car_no&"'; "

	Set rs_car = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If Not(rs_car.bof Or rs_car.eof) Then
		nCarCnt = CInt(rs_car("cnt"))
		end_msg = "이미 등록된 차량입니다."
	End If

	rs_car.Close() : Set rs_car = Nothing

	If nCarCnt > 0 Then
		Response.Write "<script type='text/javascript'>"
		Response.write "	alert('"&end_msg&"');"
		Response.write "	history.go(-1);"
		Response.write "</script>"
	End If

	'차량 정보 입력
	objBuilder.Append "INSERT INTO car_info(car_no, car_name, car_year, oil_kind, insurance_amt, "
	objBuilder.Append "car_owner,buy_gubun,rental_company,car_reg_date,car_use_dept, "
	objBuilder.Append "car_company,car_use,owner_emp_no,owner_emp_name,start_date, "
	objBuilder.Append "last_km,last_check_date,car_status,car_comment,reg_emp_name,"
	objBuilder.Append "reg_date)VALUES("
	objBuilder.Append "'"&car_no&"','"&car_name&"','"&car_year&"','"&oil_kind&"',0, "
	objBuilder.Append "'"&car_owner&"','"&buy_gubun&"','"&rental_company&"','"&car_reg_date&"','"&car_use_dept&"', "
	objBuilder.Append "'"&car_company&"','"&car_use&"','"&owner_emp_no&"','"&owner_emp_name&"','"&start_date&"', "
	objBuilder.Append "'"&last_km&"','"&last_check_date&"','"&car_status&"','"&car_comment&"','"&user_name&"', "
	objBuilder.Append "NOW()); "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'차량 운행 정보 입력
	objBuilder.Append "INSERT INTO car_drive_user(use_car_no, use_owner_emp_no, use_date, use_company, use_org_code, "
	objBuilder.Append "use_org_name, use_emp_name, use_emp_grade, use_reg_date, use_reg_user)VALUES("
	objBuilder.Append "'"&car_no&"', '"&owner_emp_no&"', '"&start_date&"', '"&emp_company&"', '"&emp_org_code&"', "
	objBuilder.Append "'"&emp_org_name&"', '"&owner_emp_name&"', '"&emp_grade&"', NOW(), '"&user_name&"'); "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "저장 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "저장 되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	parent.opener.location.reload();"
Response.Write "	self.close();"
Response.Write "</script>"
Response.End
%>
