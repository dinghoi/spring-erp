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
Dim edu_empno, edu_seq, edu_empname, owner_view, end_msg, url

edu_empno = Request.form("edu_empno")
edu_seq = Request.form("edu_seq")
edu_empname = Request.form("edu_empname")
owner_view = Request.form("owner_view")

DBConn.BeginTrans

'sql = " delete from emp_edu " & _
'			"  where edu_empno ='"&edu_empno&"' and edu_seq = '"&edu_seq&"'"
objBuilder.Append "DELETE FROM emp_edu WHERE edu_empno ='"&edu_empno&"' AND edu_seq = '"&edu_seq&"';"

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

'url = "insa_family_mg.asp?ck_sw="y"&view_condi=" + family_empno + "&ck_sw= y&view_condi="+view_condi+"&condi="+ condi
'url = "insa_edu_mg.asp?ck_sw=y&view_condi=" + edu_empno + "&condi="+ edu_empname
url = "/insa/insa_edu_mg.asp?owner_view="&owner_view

If owner_view = "C" Then
	url = url&"&view_condi="&edu_empname
else
	url = url&"&view_condi="&edu_empno
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('"&url&"');"
Response.Write "</script>"
Response.End
%>
