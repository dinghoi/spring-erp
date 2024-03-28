<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'on Error resume next

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
Dim org_company, end_month, end_yn, cost_year, cost_month
Dim from_date, end_date, to_date
Dim rs, end_msg, deptName

org_company = Request("org_company")
'saupbu = "사업부외나머지"
deptName = "사업부외나머지"
end_month = Request("end_month")
end_yn = Request("end_yn")

cost_year = Mid(end_month, 1, 4)
cost_month = Mid(end_month, 5)

from_date = Mid(end_month, 1, 4) & "-" & Mid(end_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('마감 취소중!!!');"
Response.Write "</script>"

DBConn.BeginTrans

'야특근 마감
'sql = "Update overtime set end_yn='C' where work_date >= '"&from_date&"' and work_date <= '"&to_date&"' and (saupbu ='')"
objBuilder.Append "UPDATE overtime SET end_yn='C' WHERE work_date >= '"&from_date&"' AND work_date <= '"&to_date&"' AND bonbu ='' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'일반비용
'sql = "Update general_cost set end_yn='C' where (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and (saupbu ='')"
objBuilder.Append "UPDATE general_cost SET end_yn='C' WHERE (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') AND bonbu ='' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'교통비
'sql = "Update transit_cost set end_yn='C' where (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and (saupbu ='')"
objBuilder.Append "UPDATE transit_cost SET end_yn='C' WHERE (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"') AND bonbu ='' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "Update cost_end set end_yn='C',batch_yn='N',bonbu_yn='N',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where end_month = '"&end_month& _
'"' and saupbu = '"&saupbu&"'"
objBuilder.Append "UPDATE cost_end SET end_yn='C', batch_yn='N', bonbu_yn='N', mod_id='"&user_id&"', mod_name = '"&user_name&"', mod_date = NOW() "
objBuilder.Append "WHERE end_month = '"&end_month&"' AND saupbu = '사업부외나머지' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "update org_cost set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (saupbu ='')"
objBuilder.Append "UPDATE org_cost SET cost_amt_"&cost_month&"= '0' WHERE cost_year ='"&cost_year&"' AND bonbu ='' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 상주비용 취소
'sql = "Update cost_end set end_yn='C',batch_yn='N',bonbu_yn='N',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where end_month = '"&end_month&"' and saupbu = '상주비용' "
objBuilder.Append "UPDATE cost_end SET end_yn='C', batch_yn='N', bonbu_yn='N', mod_id='"&user_id&"', mod_name='"&user_name&"', mod_date=NOW() "
objBuilder.Append "WHERE end_month = '"&end_month&"' AND saupbu = '상주비용' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "update company_cost set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
objBuilder.Append "UPDATE company_cost SET cost_amt_"&cost_month&"= '0' WHERE cost_year ='"&cost_year&"' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
objBuilder.Append "UPDATE company_profit_loss SET cost_amt_"&cost_month&"= '0' WHERE cost_year ='"&cost_year&"' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
objBuilder.Append "UPDATE company_profit_loss SET cost_amt_"&cost_month&"= '0' WHERE cost_year ='"&cost_year&"' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 공통비 배분 취소
'sql = "Update cost_end set end_yn='C',batch_yn='N',bonbu_yn='N',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where end_month = '"&end_month&"' and saupbu = '공통비/직접비배분'"
objBuilder.Append "UPDATE cost_end SET end_yn='C', batch_yn='N', bonbu_yn='N', mod_id='"&user_id&"', mod_name='"&user_name&"', mod_date=NOW() "
objBuilder.Append "WHERE end_month = '"&end_month&"' AND saupbu = '공통비/직접비배분' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "delete from company_as where as_month ='"&end_month&"'"
objBuilder.Append "DELETE FROM company_as WHERE as_month ='"&end_month&"' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "delete from company_asunit where as_month ='"&end_month&"'" ' AS 표준단가
objBuilder.Append "DELETE FROM company_asunit WHERE as_month ='"&end_month&"' "
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "delete from management_cost where cost_month ='"&end_month&"'"
objBuilder.Append "DELETE FROM management_cost WHERE cost_month ='"&end_month&"'"
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "처리중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "마감이 취소되었습니다...."
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('cost_end_mg.asp');"
Response.Write "</script>"
Response.End

DBConn.Close() : Set DBConn = Nothing
%>


