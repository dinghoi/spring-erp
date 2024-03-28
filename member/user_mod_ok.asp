<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
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
Dim pass, mod_pass, hp, car_yn, old_car_yn
Dim car_no, old_car_no, car_name, car_owner, oil_kind
Dim curr_date, carRS, end_msg

pass = Request.Form("pass")
mod_pass = Request.Form("mod_re_pass")
hp = Request.Form("hp")
car_yn = Request.Form("car_yn")
old_car_yn = Request.Form("Formld_car_yn")
car_no = Request.Form("car_no")
old_car_no = Request.Form("old_car_no")
car_name = Request.Form("car_name")
car_owner = Request.Form("car_owner")
oil_kind = Request.Form("oil_kind")
curr_date = Mid(Now(), 1, 10)

DBConn.BeginTrans

If mod_pass = "" Then
	objBuilder.Append "UPDATE memb SET "
	objBuilder.Append "	hp = '"&hp&"', car_yn = '"&car_yn&"', mod_id = '"&user_id&"', mod_date = NOW() "
	objBuilder.Append "WHERE user_id = '"&user_id&"' "
Else
	objBuilder.Append "UPDATE memb SET "
	objBuilder.Append "	pass = '"&mod_pass&"', hp = '"&hp&"', car_yn = '"&car_yn&"', mod_id = '"&user_id&"', mod_date = NOW() "
	objBuilder.Append "WHERE user_id='"&user_id&"' "
End If

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If car_yn = "Y" Then
	objBuilder.Append "SELECT owner_emp_no FROM car_info WHERE owner_emp_no ='" & user_id & "'"

	Set carRs = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If carRs.EOF Or carRs.BOF Then
		objBuilder.Append "INSERT INTO car_info(car_no, car_name, oil_kind, car_owner, buy_gubun, "
		objBuilder.Append "car_reg_date, owner_emp_no, owner_emp_name, start_date, last_km, "
		objBuilder.Append "reg_emp_no, reg_emp_name, reg_date, insurance_amt)VALUES("
		objBuilder.Append "'"&car_no&"', '"&car_name&"', '"&oil_kind&"', '개인', '구매', "
		objBuilder.Append "'"&curr_date&"', '"&user_id&"', '"&user_name&"', '"&curr_date&"', 0, "
		objBuilder.Append "'"&user_id&"', '"&user_name&"', NOW(), 0) "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Else
		If old_car_no = car_no Then
			objBuilder.Append "UPDATE car_info SET "
			objBuilder.Append "	car_name = '"&car_name&"', oil_kind = '"&oil_kind&"', mod_emp_no = '"&user_id&"', "
			objBuilder.Append "	mod_emp_name = '"&user_name&"', mod_date = NOW() "
			objBuilder.Append "WHERE owner_emp_no = '"&user_id&"' "

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
		Else
			objBuilder.Append "DELETE FROM car_info WHERE owner_emp_no = '"&user_id&"' "

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			objBuilder.Append "INSERT INTO car_info(car_no, car_name, oil_kind, car_owner, buy_gubun, "
			objBuilder.Append "car_reg_date, owner_emp_no, start_date, last_km, reg_emp_no, "
			objBuilder.Append "reg_emp_name, reg_date, insurance_amt)VALUES("
			objBuilder.Append "	'"&car_no&"','"&car_name&"','"&oil_kind&"','개인','구매', "
			objBuilder.Append "'"&curr_date&"','"&user_id&"','"&curr_date&"',0,'"&user_id&"', "
			objBuilder.Append "'"&user_name&"', NOW(), 0) "

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
		End If
	End If

	carRs.Close() : Set carRs = Nothing
End If

If car_yn = "N" And old_car_yn = "Y" And car_owner = "개인" Then
	objBuilder.Append "DELETE FROM car_info WHERE owner_emp_no = '"&user_id&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "변경 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "정상적으로 변경되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	parent.opener.location.reload();"
Response.Write "	self.close();"
Response.Write "</script>"

Response.End
%>

