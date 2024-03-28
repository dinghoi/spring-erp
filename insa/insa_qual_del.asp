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
Dim qual_empno, qual_seq, qual_name, owner_view, url, end_msg

qual_empno = Request.Form("qual_empno")
qual_seq = Request.Form("qual_seq")
qual_name = Request.Form("qual_name")
owner_view = Request.Form("owner_view")

DBConn.BeginTrans

objBuilder.Append "DELETE FROM emp_qual WHERE qual_empno ='"&qual_empno&"' AND qual_seq = '"&qual_seq&"';"

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

url = "/insa/insa_qual_mg.asp?owner_view="&owner_view

If owner_view = "C" Then
	url = url&"&view_condi="&qual_name
Else
	url = url&"&view_condi="&qual_empno
End If

Response.write "<script type='text/javascript'>"
Response.write "	alert('"&end_msg&"');"
Response.write "	location.replace('"&url&"');"
Response.write "</script>"
Response.End
%>
