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
Dim u_type, edu_seq, edu_empno, edu_start_date, edu_end_date
Dim edu_name, edu_office, edu_finish_no, edu_pay, edu_comment
Dim max_seq, rsSeq, end_msg

u_type = request.form("u_type")
edu_seq = request.form("edu_seq")
edu_empno = request.form("edu_empno")
edu_start_date = request.form("edu_start_date")
edu_end_date = request.form("edu_end_date")
edu_name = request.form("edu_name")
edu_office = request.form("edu_office")
edu_finish_no = request.form("edu_finish_no")
edu_pay = 0
'edu_pay = request.form("edu_pay")
edu_comment = request.form("edu_comment")
'edu_reg_date = request.form("edu_reg_date")

DBConn.BeginTrans

If u_type = "U" Then
	objBuilder.Append "UPDATE emp_edu SET "
	objBuilder.Append "	edu_name='"&edu_name&"',edu_office='"&edu_office&"',edu_finish_no='"&edu_finish_no&"',"
	objBuilder.Append "	edu_start_date='"&edu_start_date&"',edu_end_date='"&edu_end_date&"',edu_comment='"&edu_comment&"',"
	objBuilder.Append "	edu_mod_date=NOW(),edu_mod_user='"&user_name&"' "
	objBuilder.Append "WHERE edu_empno ='"&edu_empno&"' AND edu_seq = '"&edu_seq&"';"
Else
	objBuilder.Append "SELECT MAX(edu_seq) AS 'max_seq' FROM emp_edu WHERE edu_empno='"&edu_empno&"';"

	Set rsSeq = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If f_toString(rsSeq("max_seq"), "") = "" Then
		edu_seq = "001"
	Else
		max_seq = "00"&CStr((Int(rsSeq("max_seq")) + 1))
		edu_seq = Right(max_seq, 3)
	End If
	rsSeq.Close() : Set rsSeq = Nothing

	objBuilder.Append "INSERT INTO emp_edu(edu_empno,edu_seq,edu_name,edu_office,edu_finish_no,"
	objBuilder.Append "edu_start_date,edu_end_date,edu_pay,edu_comment,edu_reg_date,edu_reg_user)"
	objBuilder.Append "VALUES('"&edu_empno&"','"&edu_seq&"','"&edu_name&"','"&edu_office&"','"&edu_finish_no&"',"
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
