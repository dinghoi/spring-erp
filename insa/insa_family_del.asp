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
Dim family_empno, family_seq, family_name, owner_view, end_msg, url

family_empno = Request.Form("family_empno")
family_seq = Request.Form("family_seq")
family_name = Request.Form("family_name")
owner_view = Request.Form("owner_view")

DBConn.BeginTrans

objBuilder.Append "DELETE FROM emp_family "
objBuilder.Append "WHERE family_empno = '"&family_empno&"' AND family_seq = '"&family_seq&"';"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "���� �� Error�� �߻��Ͽ����ϴ�."
Else
	DBConn.CommitTrans
	end_msg = "���������� �����Ǿ����ϴ�."
End If

DBConn.Close() : Set DBConn = Nothing

url = "/insa/insa_family_mg.asp?owner_view="&owner_view&"&view_condi="

If owner_view = "C" Then
	url = url&family_name
Else
	url = url&family_empno
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
'response.write "	location.replace('insa_family_mg.asp');"
Response.Write "	location.replace('"&url&"');"
Response.Write "</script>"
Response.End
%>
