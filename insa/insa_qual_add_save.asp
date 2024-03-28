<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim u_type, qual_seq, qual_empno, qual_type, qual_grade, qual_pass_date
Dim qual_org, qual_no, qual_passport, qual_pay_id, max_seq, rsQual, end_msg

u_type = Request.Form("u_type")
qual_seq = Request.Form("qual_seq")
qual_empno = Request.Form("qual_empno")
qual_type = Request.Form("qual_type")
qual_grade = Request.Form("qual_grade")
qual_pass_date = Request.Form("qual_pass_date")
qual_org = Request.Form("qual_org")
qual_no = Request.Form("qual_no")
qual_passport = Request.Form("qual_passport")
qual_pay_id = Request.Form("qual_pay_id")

DBConn.BeginTrans

If u_type = "U" Then
	'sql = "update emp_qual set qual_type='"&qual_type&"',qual_grade='"&qual_grade&"',qual_pass_date='"&qual_pass_date&"',qual_org='"&qual_org&"',qual_no='"&qual_no&"',qual_passport='"&qual_passport&"',qual_pay_id='"&qual_pay_id&"',qual_mod_date=now(),qual_mod_user='"&emp_user&"' where qual_empno ='"&qual_empno&"' and qual_seq = '"&qual_seq&"'"
	objBuilder.Append "UPDATE emp_qual SET "
	objBuilder.Append "	qual_type='"&qual_type&"',qual_grade='"&qual_grade&"',qual_pass_date='"&qual_pass_date&"',"
	objBuilder.Append "	qual_org='"&qual_org&"',qual_no='"&qual_no&"',qual_passport='"&qual_passport&"',"
	objBuilder.Append "	qual_pay_id='"&qual_pay_id&"',qual_mod_date=NOW(),qual_mod_user='"&user_name&"' "
	objBuilder.Append "WHERE qual_empno ='"&qual_empno&"' AND qual_seq = '"&qual_seq&"';"

Else
	objBuilder.Append "SELECT MAX(qual_seq) AS 'max_seq' FROM emp_qual WHERE qual_empno='"&qual_empno&"';"

	Set rsQual = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If f_toString(rsQual("max_seq"), "") = ""  Then
		qual_seq = "001"
	Else
		max_seq = "00"&CStr((Int(rsQual("max_seq")) + 1))
		qual_seq = Right(max_seq,3)
	End If

	'sql = "insert into emp_qual(qual_empno,qual_seq,qual_type,qual_grade,qual_pass_date,qual_org,qual_no,qual_passport,qual_pay_id,qual_reg_date,qual_reg_user) values "
	'sql = sql +	" ('"&qual_empno&"','"&qual_seq&"','"&qual_type&"','"&qual_grade&"','"&qual_pass_date&"','"&qual_org&"','"&qual_no&"','"&qual_passport&"','"&qual_pay_id&"',now(),'"&emp_user&"')"

	objBuilder.Append "INSERT INTO emp_qual(qual_empno,qual_seq,qual_type,qual_grade,qual_pass_date,"
	objBuilder.Append "qual_org,qual_no,qual_passport,qual_pay_id,qual_reg_date,qual_reg_user)"
	objBuilder.Append "VALUES('"&qual_empno&"','"&qual_seq&"','"&qual_type&"','"&qual_grade&"','"&qual_pass_date&"',"
	objBuilder.Append "'"&qual_org&"','"&qual_no&"','"&qual_passport&"','"&qual_pay_id&"',NOW(),'"&user_name&"');"
End If

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	dbconn.RollbackTrans
	end_msg = "등록 중 Error가 발생하였습니다."
Else
	dbconn.CommitTrans
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
