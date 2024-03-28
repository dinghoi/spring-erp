<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### 작업 내역
'===================================================
' 허정호_20210721 :
'	- 신규 페이지 작성 및 코드 정리
'	- 보험은 갱신 개념으로 추가만 가능하게 작성, 별도 관리 페이지나 nkp에서 관리하지 않음(문의:인사 이윤정 과장)

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
'on Error resume next

Dim car_no, car_name, car_year, car_reg_date, ins_car_no
Dim ins_date, ins_old_date, ins_amount, ins_company, ins_last_date
Dim ins_man1, ins_man2, ins_object, ins_self, ins_injury
Dim ins_self_car, ins_age, ins_comment, ins_contract_yn
Dim ins_scramble, end_msg


'u_type = request.form("u_type")

car_no = Request.Form("car_no")
car_name = Request.Form("car_name")
car_year = Request.Form("car_year")
car_reg_date = Request.Form("car_reg_date")
ins_car_no = Request.Form("car_no")
ins_date = Request.Form("ins_date")
ins_old_date = Request.Form("ins_old_date")
ins_amount = Int(Request.Form("ins_amount"))
ins_company = Request.Form("ins_company")
ins_last_date = Request.Form("ins_last_date")
ins_man1 = Request.Form("ins_man1")
ins_man2 = Request.Form("ins_man2")
ins_object = Request.Form("ins_object")
ins_self = Request.Form("ins_self")
ins_injury = Request.Form("ins_injury")
ins_self_car = Request.Form("ins_self_car")
ins_age = Request.Form("ins_age")

ins_comment = Request.Form("ins_comment")
ins_contract_yn = Request.Form("ins_contract_yn")

If ins_contract_yn = "N" Then
	ins_comment = "필요시 제안사에서 운영"
Else
	ins_comment = ""
End If

ins_scramble = Request.Form("ins_scramble")

DBConn.BeginTrans

'emp_user = request.cookies("nkpmg_user")("coo_user_name")

objBuilder.Append "INSERT INTO car_insurance(ins_car_no, ins_date, ins_amount, ins_company, ins_last_date, "
objBuilder.Append "ins_man1, ins_man2, ins_object, ins_self, ins_injury, "
objBuilder.Append "ins_self_car, ins_age, ins_comment, ins_contract_yn, ins_scramble, "
objBuilder.Append "ins_reg_date,ins_reg_user)VALUES("
objBuilder.Append "'"&ins_car_no&"','"&ins_date&"','"&ins_amount&"','"&ins_company&"','"&ins_last_date&"', "
objBuilder.Append "'"&ins_man1&"','"&ins_man2&"','"&ins_object&"','"&ins_self&"','"&ins_injury&"', "
objBuilder.Append "'"&ins_self_car&"','"&ins_age&"','"&ins_comment&"','"&ins_contract_yn&"','"&ins_scramble&"', "
objBuilder.Append "NOW(),'"&user_name&"');"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "UPDATE car_info SET "
objBuilder.Append "	insurance_company='"&ins_company&"', "
objBuilder.Append "	insurance_date ='"&ins_last_date&"', "
objBuilder.Append "	insurance_amt ='"&ins_amount&"', "
objBuilder.Append "	mod_emp_name='"&user_name&"', "
objBuilder.Append "	mod_date = NOW() "
objBuilder.Append "WHERE car_no = '"&ins_car_no&"'; "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "자장중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "저장되었습니다."
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	parent.opener.location.reload();"
Response.Write "	self.close() ;"
Response.Write "</script>"
Response.End

dbconn.Close() : Set dbconn = Nothing
%>
