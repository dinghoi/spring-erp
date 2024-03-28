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
Dim end_month, cost_year, cost_month, from_date
Dim end_date, to_date, end_msg

end_month = Request("end_month")

cost_year = Mid(end_month, 1, 4)
cost_month = Mid(end_month, 5)

from_date = Mid(end_month, 1, 4) & "-" & Mid(end_month, 5, 2) & "-01"

end_date = DateValue(from_date)
end_date = DateAdd("m",1,from_date)
to_date = CStr(DateAdd("d", -1, end_date))

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('마감 취소중!!!');"
Response.Write "</script>"

DBConn.BeginTrans

'sql = "Update cost_end set end_yn='C',batch_yn='N',bonbu_yn='N',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where end_month = '"&end_month&"' and saupbu = '상주비용'"
objBuilder.Append "UPDATE cost_end SET "
objBuilder.Append "	end_yn='C', batch_yn = 'N', bonbu_yn = 'N', mod_id = '"&user_id&"', "
objBuilder.Append "	mod_name = '"&user_name&"', mod_date = NOW() "
objBuilder.Append "WHERE end_month = '"&end_month&"' AND saupbu = '상주비용' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "update company_cost set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
objBuilder.Append "UPDATE company_cost SET cost_amt_"&cost_month&" = '0' WHERE cost_year ='"&cost_year&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
objBuilder.Append "UPDATE company_profit_loss SET cost_amt_"&cost_month&" = '0' WHERE cost_year ='"&cost_year&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
objBuilder.Append "UPDATE saupbu_profit_loss SET cost_amt_"&cost_month&" = '0' WHERE cost_year ='"&cost_year&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "처리중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "마감이 취소되었습니다."
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('/cost/cost_end_mg.asp');"
Response.Write "</script>"
Response.End

DBConn.Close() : Set DBConn = Nothing
%>


