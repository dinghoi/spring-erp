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
Dim view_condi, rs_emp, end_msg, emp_person2

view_condi = Request.Form("view_condi1")

objBuilder.Append "SELECT emp_person2 FROM emp_master "
objBuilder.Append "WHERE emp_no = '"&view_condi&"';"

Set rs_emp = DbConn.Execute(objBuilder.ToString())
objBuilder.Clear()

emp_person2 = rs_emp("emp_person2")

rs_emp.Close() : Set rs_emp = Nothing

If f_toString(emp_person2, "") = "" Then
	 emp_person2 = view_condi
End If

DBConn.BeginTrans

'sql = "Update memb set pass='"&emp_person2&"',mod_id ='"&user_id&"',mod_date=now() where user_id='"&view_condi&"'"
objBuilder.Append "UPDATE memb SET "
objBuilder.Append "	pass='"&emp_person2&"', mod_id ='"&user_id&"', mod_date=NOW() "
objBuilder.Append "WHERE user_id='"&view_condi&"'"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "변경 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "정상적으로 변경되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
response.write "	alert('"&end_msg&"');"
response.write "	parent.opener.location.reload();"
response.write "	self.close() ;"
response.write "</script>"
Response.End
%>
