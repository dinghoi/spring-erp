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
Dim u_type, etc_code, etc_typew, type_name, etc_group, group_name
Dim used_sw, emp_tax_id, etc_type, etc_name, end_msg

u_type = Request.Form("u_type")
etc_code = Request.Form("etc_code")
etc_type = Request.Form("etc_type")
type_name = Request.Form("type_name")
etc_name = Request.Form("etc_name")
etc_group = Request.Form("etc_group")
group_name = Request.Form("group_name")
used_sw = Request.Form("used_sw")
emp_tax_id = Request.Form("emp_tax_id")

DBConn.BeginTrans

If u_type = "U" Then
	objBuilder.Append "UPDATE emp_etc_code SET "
	objBuilder.Append "	emp_etc_name='"&etc_name&"',emp_etc_group='"&etc_group&"',emp_group_name ='"&group_name&"',"
	objBuilder.Append "	emp_used_sw='"&used_sw&"',emp_tax_id='"&emp_tax_id&"' "
	objBuilder.Append "WHERE emp_etc_code = '"&etc_code&"';"
Else
	objBuilder.Append "INSERT INTO emp_etc_code(emp_etc_code,emp_etc_type,emp_type_name,emp_etc_name,emp_etc_group,"
	objBuilder.Append "emp_group_name,emp_mg_group,emp_used_sw,emp_tax_id)"
	objBuilder.Append "VALUES('"&etc_code&"','"&etc_type&"','"&type_name&"','"&etc_name&"','"&etc_group&"',"
	objBuilder.Append "'"&group_name&"','"&mg_group&"','"&used_sw&"','"&emp_tax_id&"');"
End If

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.Number <> "0" Then
	DBConn.RollbackTrans
	end_msg = "등록 중 오류가 발생했습니다."
Else
	DBConn.CommitTrans
	end_msg = "정상적으로 등록되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.write "	alert('"&end_msg&"');"
Response.Write "	location.href='/pay/insa_pay_code_mg.asp?emp_etc_type="&etc_type&"';"
Response.write "</script>"
Response.End
%>
