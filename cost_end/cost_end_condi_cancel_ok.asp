<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
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
'on Error resume next

Dim from_month, to_month, from_date, to_date
Dim cost_year, cost_month, rsEnd, yyyymm, end_msg

from_month = f_Request("from_month")
to_month = f_Request("to_month")

from_date = Mid(from_month, 1, 4) & "-" & Mid(from_month, 5, 2)
to_date = Mid(to_month, 1, 4) & "-" & Mid(to_month, 5, 2)

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('마감 취소중!!!');"
Response.Write "</script>"

DBConn.BeginTrans

' 조직별 비용 CLEAR
for yyyymm = from_month to to_month
	cost_year = mid(yyyymm,1,4)
	cost_month = cstr(mid(yyyymm,5))

	objBuilder.Append "UPDATE org_cost SET cost_amt_"&cost_month&"= '0' WHERE cost_year ='"&cost_year&"';"
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	objBuilder.Append "UPDATE company_cost SET cost_amt_"&cost_month&"= '0' WHERE cost_year ='"&cost_year&"';"
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	objBuilder.Append "UPDATE company_profit_loss SET cost_amt_"&cost_month&"= '0' WHERE cost_year ='"&cost_year&"';"
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	objBuilder.Append "UPDATE saupbu_profit_loss SET cost_amt_"&cost_month&"= '0' WHERE cost_year ='"&cost_year&"';"
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
Next

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "처리중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans

	objBuilder.Append "CALL USP_COST_CANCEL_UPDATE('"&from_date&"', '"&to_date&"', '"&from_month&"', '"&to_month&"');"
	Set rsEnd = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsEnd(0) = "0" Then
		end_msg = "마감이 취소되었습니다."
	End If

	rsEnd.Close() : Set rsEnd = Nothing
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('/cost_end/cost_end_mg.asp');"
Response.write "</script>"
Response.End

DBConn.Close() : Set DBConn = Nothing
%>
