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
Dim u_type, career_seq, career_empno, career_join_date, career_end_date, career_office
Dim career_dept, career_position, career_task, rsMax, end_msg, max_seq

u_type = request.form("u_type")
career_seq = request.form("career_seq")
career_empno = request.form("career_empno")

career_join_date = request.form("career_join_date")
career_end_date = request.form("career_end_date")
career_office = request.form("career_office")
career_dept = request.form("career_dept")
career_position = request.form("career_position")
career_task = request.form("career_task")


DBConn.BeginTrans

If u_type = "U" Then
	objBuilder.Append "UPDATE emp_career SET "
	objBuilder.Append "	career_join_date='"&career_join_date&"',career_end_date='"&career_end_date&"',career_office='"&career_office&"', "
	objBuilder.Append "	career_dept='"&career_dept&"',career_position='"&career_position&"',career_task='"&career_task&"', "
	objBuilder.Append "	career_mod_date=NOW(),career_mod_user='"&user_name&"' "
	objBuilder.Append "WHERE career_empno ='"&career_empno&"' AND career_seq = '"&career_seq&"';"
Else
	objBuilder.Append "SELECT MAX(career_seq) AS 'max_seq' FROM emp_career "
	objBuilder.Append "WHERE career_empno='"&career_empno&"';"

	Set rsMax = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If f_toString(rsMax("max_seq"), "") = ""  Then
		career_seq = "001"
	Else
		max_seq = "00"&CStr((Int(rsMax("max_seq")) + 1))
		career_seq = Right(max_seq, 3)
	End If

	objBuilder.Append "INSERT INTO emp_career(career_empno,career_seq,career_join_date,career_end_date,career_office,"
	objBuilder.Append "career_dept,career_position,career_task,career_reg_date,career_reg_user)VALUES("
	objBuilder.Append "'"&career_empno&"','"&career_seq&"','"&career_join_date&"','"&career_end_date&"','"&career_office&"',"
	objBuilder.Append "'"&career_dept&"','"&career_position&"','"&career_task&"',NOW(),'"&user_name&"');"
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
