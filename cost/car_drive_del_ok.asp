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
Dim u_type, mg_ce_id, run_date, old_date, run_seq
Dim end_msg

u_type = f_Request("u_type")
mg_ce_id = f_Request("mg_ce_id")
run_date = f_Request("run_date")
old_date = f_Request("old_date")
run_seq = f_Request("run_seq")

DBConn.BeginTrans

'sql = "delete from transit_cost where run_date ='"&old_date&"' and mg_ce_id='"&mg_ce_id&"' and run_seq= '"&run_seq&"'"
objBuilder.Append "DELETE FROM transit_cost "
objBuilder.Append "WHERE run_date ='"&old_date&"' AND mg_ce_id='"&mg_ce_id&"' AND run_seq= '"&run_seq&"';"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "삭제 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "정상적으로 삭제되었습니다."
End If

DBConn.Close():Set DBConn = Nothing

Response.write "<script type='text/javascript'>"
Response.write "	alert('"&end_msg&"');"
Response.write "	self.opener.location.reload();"
Response.write "	window.close();"
Response.write "</script>"
Response.End
%>
