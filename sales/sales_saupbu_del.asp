<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
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
Dim approve_no, end_msg

'on Error resume next

approve_no = Request.Form("approve_no")

DBConn.BeginTrans

objBuilder.Append "DELETE FROM saupbu_sales "
objBuilder.Append "WHERE approve_no = '"&approve_no&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "삭제중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "삭제 처리 되었습니다."
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.write "	self.opener.location.reload();"
Response.write "	window.close();"
Response.Write "</script>"
Response.End

DBConn.Close() : Set DBConn = Nothing
%>


