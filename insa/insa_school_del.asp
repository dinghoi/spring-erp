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
Dim sch_empno, sch_seq, sch_emp_name, owner_view, end_msg
Dim url

sch_empno = Request.form("sch_empno")
sch_seq = Request.form("sch_seq")
sch_emp_name = Request.form("sch_emp_name")
owner_view = Request.form("owner_view")

DBConn.BeginTrans

'sql = " delete from emp_school " & _
'			"  where sch_empno ='"&sch_empno&"' and sch_seq = '"&sch_seq&"'"
objBuilder.Append "DELETE FROM emp_school WHERE sch_empno ='"&sch_empno&"' AND sch_seq = '"&sch_seq&"';"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "삭제 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "정상적으로 삭제되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

url = "/insa/insa_school_mg.asp?owner_view="&owner_view

If owner_view = "C" Then
	url = url&"&view_condi="&sch_emp_name
Else
	url = url&"&view_condi="&sch_empno
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
'Response.Write "	location.replace('insa_family_mg.asp');"
Response.Write "	location.replace('"&url&"');"
Response.Write "</script>"
Response.End
%>
