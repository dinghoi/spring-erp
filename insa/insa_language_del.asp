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
Dim lang_empno, lang_seq, lang_empname, owner_view
Dim end_msg, url

lang_empno = Request.Form("lang_empno")
lang_seq = Request.Form("lang_seq")
lang_empname = Request.Form("lang_empname")
owner_view = Request.Form("owner_view")

DBConn.BeginTrans

'sql = " delete from emp_language " & _
'			"  where lang_empno ='"&lang_empno&"' and lang_seq = '"&lang_seq&"'"

objBuilder.Append "DELETE FROM emp_language WHERE lang_empno ='"&lang_empno&"' AND lang_seq = '"&lang_seq&"';"

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

url = "/insa/insa_language_mg.asp?owner_view="&owner_view

If owner_view = "C" Then
	   url = url&"&view_condi="&lang_empname
Else
	   url = url&"&view_condi="&lang_empno
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('"&url&"');"
Response.Write "</script>"
Response.End
%>
