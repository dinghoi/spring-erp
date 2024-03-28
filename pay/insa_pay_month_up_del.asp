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
Dim pay_company, pay_month, end_msg, pg_url

pay_company = f_Request("pay_company1")
pay_month = f_Request("pay_month1")

DBConn.BeginTrans

objBuilder.Append "DELETE FROM pay_month_give "
objBuilder.Append "WHERE pmg_yymm = '"&pay_month&"' AND pmg_id = '1' "
objBuilder.Append "	AND pmg_company = '"&pay_company&"';"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "DELETE FROM pay_month_deduct "
objBuilder.Append "WHERE de_yymm = '"&pay_month&"' AND de_id = '1' "
objBUilder.Append "	AND de_company = '"&pay_company&"';"

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

pg_url = "/pay/insa_pay_month_up.asp?ck_sw=y&pay_company="&pay_company&"&pay_month="&pay_month

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('"&pg_url&"');"
Response.Write "</script>"
Response.End
%>
