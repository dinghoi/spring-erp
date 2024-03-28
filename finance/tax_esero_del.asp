<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
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
Dim approve_no, end_msg
Dim page, bill_id, bill_month, cost_reg_yn, end_yn, url

approve_no = Request.QueryString("t_id")

page = Request.QueryString("page")
bill_id = Request.QueryString("b_id")
bill_month = Request.QueryString("b_month")
cost_reg_yn = Request.QueryString("c_yn")
end_yn = Request.QueryString("e_yn")

DBConn.BeginTrans

'�̼��� ���Լ��ݰ�꼭 ����
objBuilder.Append "DELETE FROM tax_bill "
objBuilder.Append "WHERE approve_no = '"&approve_no&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "ó�� �� Error�� �߻��Ͽ����ϴ�."
Else
	DBConn.CommitTrans
	end_msg = "���� ó���Ǿ����ϴ�."
End If

DBConn.Close() : Set DBConn = Nothing

url = "/finance/tax_esero_mg.asp?page="&page&"&bill_id="&bill_id&"&bill_month="&bill_month&"&cost_reg_yn="&cost_reg_yn&"&end_yn="&end_yn

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('"&url&"');"
Response.Write "</script>"
Response.End
%>


