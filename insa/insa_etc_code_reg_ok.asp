<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim u_type, etc_code, etc_type, type_name, etc_name, etc_group
Dim group_name, used_sw, emp_tax_id

u_type = Request.Form("u_type")
etc_code = Request.Form("etc_code")
etc_type = Request.Form("etc_type")
type_name = Request.Form("type_name")
etc_name = Request.Form("etc_name")
etc_group = Request.Form("etc_group")
group_name = Request.Form("group_name")
used_sw = Request.Form("used_sw")

emp_tax_id = ""

If u_type = "U" Then
	objBuilder.Append "UPDATE emp_etc_code SET "
	objBuilder.Append "	emp_etc_name='"&etc_name&"', emp_etc_group='"&etc_group&"', "
	objBuilder.Append "	emp_group_name='"&group_name&"', emp_used_sw='"&used_sw&"', "
	objBuilder.Append "	emp_tax_id='"&emp_tax_id&"' "
	objBuilder.Append "WHERE emp_etc_code = '"&etc_code&"'; "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
Else
	objBuilder.Append "INSERT INTO emp_etc_code(emp_etc_code, emp_etc_type, emp_type_name, emp_etc_name, emp_etc_group, "
	objBuilder.Append "emp_group_name, emp_mg_group, emp_used_sw, emp_tax_id)VALUES("
	objBuilder.Append "'"&etc_code&"','"&etc_type&"','"&type_name&"','"&etc_name&"','"&etc_group&"',"
	objBuilder.Append "'"&group_name&"','"&mg_group&"','"&used_sw&"','"&emp_tax_id&"');"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('정상적으로 등록 완료 되었습니다.');"
'Response.Redirect "	/insa/insa_etc_code_mg.asp?emp_etc_type="&etc_type
Response.Write "	location.href = '/insa/insa_etc_code_mg.asp?etc_type="&etc_type&"' "
Response.Write "</script>"
Response.End
%>
