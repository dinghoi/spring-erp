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
Dim u_type, sch_seq, sch_empno, view_condi, sch_school_name
Dim sch_start_date, sch_end_date, sch_dept, sch_major, sch_sub_major
Dim sch_degree, sch_finish, sch_comment, rsMax, end_msg, max_seq

u_type = Request.Form("u_type")
sch_seq = Request.Form("sch_seq")
sch_empno = Request.Form("sch_empno")
view_condi = Request.Form("view_condi")

If view_condi = "1" Then
	sch_school_name = Request.Form("sch_high_name")
Else
	sch_school_name = Request.Form("sch_school_name")
End If

sch_start_date = Request.Form("sch_start_date")
sch_end_date = Request.Form("sch_end_date")
sch_dept = Request.Form("sch_dept")
sch_major = Request.Form("sch_major")
sch_sub_major = Request.Form("sch_sub_major")
sch_degree = Request.Form("sch_degree")
sch_finish = Request.Form("sch_finish")
sch_comment = view_condi

DBConn.BeginTrans

If u_type = "U" Then
	objBuilder.Append "UPDATE emp_school SET "
	objBuilder.Append "	sch_start_date='"&sch_start_date&"',sch_end_date='"&sch_end_date&"',sch_school_name='"&sch_school_name&"',"
	objBuilder.Append "	sch_dept='"&sch_dept&"',sch_major='"&sch_major&"',sch_sub_major='"&sch_sub_major&"',"
	objBuilder.Append "	sch_degree='"&sch_degree&"',sch_finish='"&sch_finish&"',sch_comment='"&sch_comment&"',"
	objBuilder.Append "	sch_mod_date=NOW(),sch_mod_user='"&user_name&"' "
	objBuilder.Append "WHERE sch_empno ='"&sch_empno&"' AND sch_seq = '"&sch_seq&"';"
Else
	objBuilder.Append "SELECT MAX(sch_seq) AS 'max_seq' FROM emp_school "
	objBuilder.Append "WHERE sch_empno = '"&sch_empno&"';"

	Set rsMax = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If f_toString(rsMax("max_seq"), "") = "" Then
		sch_seq = "001"
	Else
		max_seq = "00"&CStr((Int(rsMax("max_seq")) + 1))
		sch_seq = Right(max_seq, 3)
	End If
	rsMax.Close() : Set rsMax = Nothing

	objBuilder.Append "INSERT INTO emp_school(sch_empno,sch_seq,sch_start_date,sch_end_date,"
	objBuilder.Append "sch_school_name,sch_dept,sch_major,sch_sub_major,sch_degree,"
	objBuilder.Append "sch_finish,sch_comment,sch_reg_date,sch_reg_user)VALUES("
	objBuilder.Append "'"&sch_empno&"','"&sch_seq&"','"&sch_start_date&"','"&sch_end_date&"','"&sch_school_name&"',"
	objBuilder.Append "'"&sch_dept&"','"&sch_major&"','"&sch_sub_major&"','"&sch_degree&"','"&sch_finish&"',"
	objBuilder.Append "'"&sch_comment&"',now(),'"&user_name&"');"
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