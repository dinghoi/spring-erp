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
Dim u_type, emp_name, person_no1, person_no2, bank_name, account_no, account_holder
Dim bank_code, emp_type, emp_pay_type, rs_etc, rs_emp, emp_user, end_msg
Dim rsBank, bank_cnt

u_type = f_Request("u_type")
emp_no = f_Request("emp_no")
emp_name = f_Request("emp_name")
person_no1 = f_Request("person_no1")
person_no2 = f_Request("person_no2")
bank_name = f_Request("bank_name")
account_no = f_Request("account_no")
account_holder = f_Request("account_holder")

bank_code = ""
emp_type = ""
emp_pay_type = ""

emp_user = emp_name

objBuilder.Append "SELECT emp_etc_code FROM emp_etc_code "
objBuilder.Append "WHERE emp_etc_type = '50' AND emp_etc_name = '"&bank_name&"' "

Set rs_etc = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

bank_code = rs_etc("emp_etc_code")

rs_etc.Close() : Set rs_etc = Nothing

objBuilder.Append "SELECT emp_type, emp_pay_type "
objBuilder.Append "FROM emp_master WHERE emp_no = '"&emp_no&"' "

Set rs_emp = DbConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rs_emp.EOF Then
	 emp_type = rs_emp("emp_type")
	 emp_pay_type = rs_emp("emp_pay_type")
End If
rs_emp.Close() : Set rs_emp = Nothing

DBConn.BeginTrans

If u_type = "U" Then
	objBuilder.Append "UPDATE pay_bank_account Set "
	objBuilder.Append "	bank_code='"&bank_code&"', bank_name='"&bank_name&"', account_no='"&account_no&"', "
	objBuilder.Append "	account_holder='"&account_holder&"', mod_date= NOW(), mod_user='"&emp_user&"' "
	objBuilder.Append "WHERE emp_no ='"&emp_no&"' "

Else
	objBuilder.Append "SELECT COUNT(*) FROM pay_bank_account WHERE emp_no = '"&emp_no&"' "

	Set rsBank = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	bank_cnt = CInt(rsBank(0))

	rsBank.Close() : Set rsBank = Nothing

	If bank_cnt = 0 Then
		objBuilder.Append "INSERT INTO pay_bank_account("
		objBuilder.Append "emp_no, emp_name, person_no1, person_no2, emp_type,"
		objBuilder.Append "emp_pay_type, bank_code, bank_name, account_no, account_holder,"
		objBuilder.Append "reg_date, reg_user"
		objBuilder.Append ")VALUES("
		objBuilder.Append "'"&emp_no&"','"&emp_name&"','"&person_no1&"','"&person_no2&"','"&emp_type&"', "
		objBuilder.Append "'"&emp_pay_type&"','"&bank_code&"','"&bank_name&"','"&account_no&"','"&account_holder&"',"
		objBuilder.Append "NOW(),'"&emp_user&"')"
	Else
		Response.Write "<script type='text/javascript'>"
		Response.Write "	alert('해당 직원의 은행 정보는 이미 등록되어 있습니다.');"
		Response.Write "	self.opener.location.reload();"
		Response.Write "	window.close();"
		Response.Write "</script>"
		Response.End
	End If
End If

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "등록 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "정상적으로 등록되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	self.opener.location.reload();"
Response.Write "	window.close();"
Response.Write "</script>"
Response.End
%>
