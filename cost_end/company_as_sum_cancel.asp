<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

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
Dim end_month, cost_year, cost_month
Dim end_msg

end_month=Request("end_month")
cost_year = Mid(end_month, 1, 4)
cost_month = Mid(end_month, 5)

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('���� �����!');"
Response.Write "</script>"

DBConn.BeginTrans

'sql = "Update cost_end set end_yn='C',batch_yn='N',bonbu_yn='N',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where end_month = '"&end_month&"' and saupbu = '�����/��������'"
objBuilder.Append "UPDATE cost_end SET  "
objBuilder.Append "	end_yn = 'C', batch_yn = 'N', bonbu_yn = 'N', "
objBuilder.Append "	mod_id = '"&user_id&"', mod_name = '"&user_name&"', mod_date = NOW() "
objBuilder.Append "WHERE end_month = '"&end_month&"' AND saupbu = '�����/��������' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "delete from company_as where as_month ='"&end_month&"'"
objBuilder.Append "DELETE FROM company_as WHERE as_month = '"&end_month&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "delete from company_asunit where as_month ='"&end_month&"'" ' AS ǥ�شܰ�
objBuilder.Append "DELETE FROM company_asunit WHERE as_month ='"&end_month&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "delete from management_cost where cost_month ='"&end_month&"'"
objBuilder.Append "DELETE FROM management_cost WHERE cost_month ='"&end_month&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
objBuilder.Append "UPDATE company_profit_loss SET cost_amt_"&cost_month&" = '0' WHERE cost_year ='"&cost_year&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
objBuilder.Append "UPDATE saupbu_profit_loss SET cost_amt_"&cost_month&"= '0' WHERE cost_year ='"&cost_year&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'�ű� AS List �ش� �� ���� ���� �߰�
'objBuilder.Append "DELETE FROM as_acpt_end "
'objBuilder.Append "WHERE REPLACE(SUBSTRING(acpt_date, 1, 7), '-', '') = '"&end_month&"' "

'DBConn.Execute(objBuilder.ToString())
'objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "ó���� Error�� �߻��Ͽ����ϴ�."
Else
	DBConn.CommitTrans
	end_msg = "������ ��ҵǾ����ϴ�."
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('/cost/cost_end_mg.asp');"
Response.Write "</script>"
Response.End

DBConn.Close() : Set DBConn = Nothing
%>


