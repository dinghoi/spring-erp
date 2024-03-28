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
Dim cancel_yn, mg_ce_id, run_date, run_seq, end_msg

cancel_yn = Request.Form("cancel_yn")
mg_ce_id = Request.Form("mg_ce_id")
run_date = Request.Form("run_date")
run_seq = Request.Form("run_seq")

DBConn.BeginTrans

'sql = "update transit_cost set cancel_yn='"&cancel_yn&"',mod_id='"&user_id&"',mod_user='"&user_name&"',mod_date=now() where mg_ce_id='"&mg_ce_id&"' and run_date = '"&run_date&"' and run_seq = '"&run_seq&"'"
objBuilder.Append "UPDATE transit_cost SET "
objBuilder.Append "	cancel_yn='"&cancel_yn&"',"
objBuilder.Append "	mod_id='"&user_id&"',"
objBuilder.Append "	mod_user='"&user_name&"',"
objBuilder.Append "	mod_date=NOW() "
objBuilder.Append "WHERE mg_ce_id='"&mg_ce_id&"' AND run_date = '"&run_date&"' AND run_seq = '"&run_seq&"';"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "저장 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "저장 되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	parent.opener.location.reload();"
Response.Write "	self.close() ;"
Response.Write "</script>"
Response.End
%>
