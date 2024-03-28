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
Dim career_empno, career_seq, career_name, owner_view, end_msg
Dim url

career_empno = Request.Form("career_empno")
career_seq = Request.Form("career_seq")
career_name = Request.Form("career_name")
owner_view = Request.Form("owner_view")

DBConn.BeginTrans

objBuilder.Append "DELETE FROM emp_career "
objBuilder.Append "WHERE career_empno ='"&career_empno&"' AND career_seq = '"&career_seq&"';"

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

url = "/insa/insa_career_mg.asp?owner_view="&owner_view

If owner_view = "C" Then
	url = url&"&view_condi="&career_name
Else
	url = url&"&view_condi="&career_empno
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('"&url&"');"
Response.Write "</script>"
Response.End
%>
