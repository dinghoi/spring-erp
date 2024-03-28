<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim approve_no, emp_name, end_msg
Dim mod_id, mod_name
'on Error resume next

approve_no = Request.Form("approve_no")
saupbu = Request.Form("saupbu")
emp_name = Request.Form("emp_name")
emp_no = Request.Form("emp_no")

'수정 사용자 정보
mod_Id = user_id
mod_Name = user_name

'set dbconn = server.CreateObject("adodb.connection")
'Set rs = Server.CreateObject("ADODB.Recordset")
'Dbconn.open dbconnect

DBConn.BeginTrans

'sql = "update saupbu_sales set saupbu='"&saupbu&"', emp_name='"&emp_name&"', emp_no='"&emp_no&"', reg_id='"&user_id&"', reg_name='"&user_name&"', reg_date=now() where approve_no='"&approve_no&"' "

objBuilder.Append "UPDATE saupbu_sales SET "
objBuilder.Append "	saupbu='"&saupbu&"', "
objBuilder.Append "	emp_name='"&emp_name&"', "
objBuilder.Append "	emp_no='"&emp_no&"', "

'objBuilder.Append "	reg_id='"&user_id&"', "
'objBuilder.Append "	reg_name='"&user_name&"', "
'objBuilder.Append "	reg_date = now() "
objBuilder.Append "	mod_id = '"&mod_id&"', "
objBuilder.Append "	mod_name = '"&mod_name&"', "
objBuilder.Append "	mod_date = NOW() "

objBuilder.Append "WHERE approve_no='"&approve_no&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "변경 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "변경 되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	self.opener.location.reload();"
Response.Write "	window.close();"
Response.Write"</script>"
Response.End
%>