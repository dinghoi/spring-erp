<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim u_type, edu_seq, edu_empno, edu_start_date, edu_end_date
Dim edu_name, edu_office, edu_finish_no, edu_comment, edu_pay
Dim sqlStr, rsMax, max_seq, end_msg

u_type = Request.Form("u_type")
edu_seq = Request.Form("edu_seq")
edu_empno = Request.Form("edu_empno")
edu_start_date = Request.Form("edu_start_date")
edu_end_date = Request.Form("edu_end_date")
edu_name = Request.Form("edu_name")
edu_office = Request.Form("edu_office")
edu_finish_no = Request.Form("edu_finish_no")
edu_comment = Request.Form("edu_comment")

'edu_pay = request.form("edu_pay")
'edu_reg_date = request.form("edu_reg_date")

edu_pay = 0

DBConn.BeginTrans

'emp_user = request.cookies("nkpmg_user")("coo_user_name")

If u_type = "U" Then
	'sql = "update emp_edu set edu_name='"&edu_name&"',edu_office='"&edu_office&"',edu_finish_no='"&edu_finish_no&"',edu_start_date='"&edu_start_date&"',edu_end_date='"&edu_end_date&"',edu_comment='"&edu_comment&"',edu_mod_date=now(),edu_mod_user='"&emp_user&"' where edu_empno ='"&edu_empno&"' and edu_seq = '"&edu_seq&"'"
	objBuilder.Append "UPDATE emp_edu SET "
	objBuilder.Append "	edu_name='"&edu_name&"',edu_office='"&edu_office&"',edu_finish_no='"&edu_finish_no&"',"
	objBuilder.Append "	edu_start_date='"&edu_start_date&"',edu_end_date='"&edu_end_date&"',edu_comment='"&edu_comment&"',"
	objBuilder.Append "	edu_mod_date=NOW(),edu_mod_user='"&user_name&"' "
	objBuilder.Append "WHERE edu_empno ='"&edu_empno&"' AND edu_seq = '"&edu_seq&"';"
Else
	sqlStr = "SELECT MAX(edu_seq) AS 'max_seq' FROM emp_edu WHERE edu_empno = '"&edu_empno&"';"
	Set rsMax = DBConn.execute(sqlStr)

	If f_toString(rsMax("max_seq"), "") = "" Then
		edu_seq = "001"
	Else
		max_seq = "00"&CStr((Int(rsMax("max_seq")) + 1))
		edu_seq = Right(max_seq,3)
	End If

	'sql = "insert into emp_edu (edu_empno,edu_seq,edu_name,edu_office,edu_finish_no,edu_start_date,edu_end_date,edu_pay,edu_comment,edu_reg_date,edu_reg_user) values "
	'sql = sql +	" ('"&edu_empno&"','"&edu_seq&"','"&edu_name&"','"&edu_office&"','"&edu_finish_no&"','"&edu_start_date&"','"&edu_end_date&"','"&edu_pay&"','"&edu_comment&"',now(),'"&emp_user&"')"
	objBuilder.Append "insert into emp_edu (edu_empno,edu_seq,edu_name,edu_office,edu_finish_no,"
	objBuilder.Append "edu_start_date,edu_end_date,edu_pay,edu_comment,edu_reg_date,edu_reg_user)VALUES("
	objBuilder.Append "'"&edu_empno&"','"&edu_seq&"','"&edu_name&"','"&edu_office&"','"&edu_finish_no&"',"
	objBuilder.Append "'"&edu_start_date&"','"&edu_end_date&"','"&edu_pay&"','"&edu_comment&"',NOW(),'"&user_name&"');"
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
