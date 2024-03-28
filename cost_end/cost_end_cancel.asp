<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'On Error Resume Next

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
Dim from_date, end_date, to_date, end_msg
Dim deptName
Dim rsOverTimeList, rsGeneralList, rsTransitList

org_company = Request("org_company")
'saupbu = Request("saupbu")
deptName = Request("saupbu")
end_month = Request("end_month")
end_yn = Request("end_yn")

cost_year = Mid(end_month, 1, 4)
cost_month = Mid(end_month, 5)

from_date = Mid(end_month, 1, 4) & "-" & Mid(end_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

Response.Write "<script language='javascript'>"
Response.Write "	alert('마감 취소중!!!');"
Response.Write "</script>"

DBConn.BeginTrans

'야특근 마감
objBuilder.Append "SELECT mg_ce_id, work_date "
objBuilder.Append "FROM overtime "
objBuilder.Append "WHERE (work_date >= '"&from_date&"' AND work_date <= '"&to_date&"') "
objBuilder.Append "AND bonbu ='"&deptName&"' "

Set rsOverTimeList = Server.CreateObject("ADODB.RecordSet")
rsOverTimeList.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsOverTimeList.EOF
	objBuilder.Append "UPDATE overtime SET "
	objBuilder.Append "	end_yn='C' "
	objBuilder.Append "WHERE work_date = '"&rsOverTimeList("work_date")&"' "
	objBuilder.Append "	AND mg_ce_id = '"&rsOverTimeList("mg_ce_id")&"'; "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsOverTimeList.MoveNext()
Loop
rsOverTimeList.Close() : Set rsOverTimeList = Nothing

'일반비용
objBuilder.Append "SELECT slip_seq, slip_date "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE (slip_date >= '"&from_date&"' AND slip_date <= '"&to_date&"')"
objBuilder.Append "	AND bonbu ='"&deptName&"' "

Set rsGeneralList = Server.CreateObject("ADODB.RecordSet")
rsGeneralList.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsGeneralList.EOF
	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	end_yn='C' "
	objBuilder.Append "WHERE slip_date = '"&rsGeneralList("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsGeneralList("slip_seq")&"'; "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsGeneralList.MoveNext()
Loop
rsGeneralList.Close() : Set rsGeneralList = Nothing

'교통비
'sql = "select * from transit_cost where (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and saupbu ='"&saupbu&"'"
objBuilder.Append "SELECT mg_ce_id, run_seq, run_date "
objBuilder.Append "FROM transit_cost "
objBuilder.Append "WHERE (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"')"
objBuilder.Append "	AND bonbu ='"&deptName&"' "

Set rsTransitList = Server.CreateObject("ADODB.RecordSet")
rsTransitList.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTransitList.EOF
	objBuilder.Append "UPDATE transit_cost SET "
	objBuilder.Append "	end_yn='C' "
	objBuilder.Append "WHERE run_date = '"&rsTransitList("run_date")&"' "
	objBuilder.Append "	AND mg_ce_id = '"&rsTransitList("mg_ce_id")&"' "
	objBuilder.Append "	AND run_seq ='"&rsTransitList("run_seq")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsTransitList.MoveNext()
Loop

rsTransitList.Close() : Set rsTransitList = Nothing

'비용 마감 처리 취소
objBuilder.Append "UPDATE cost_end SET "
objBuilder.Append "	end_yn = 'C', batch_yn = 'N', bonbu_yn = 'N', "
objBuilder.Append "	mod_id = '"&user_id&"', mod_name = '"&user_name&"', mod_date = NOW() "
objBuilder.Append "WHERE end_month = '"&end_month&"' "
objBuilder.Append "	AND saupbu = '"&deptName&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'일반경비 마감 처리 취소
objBuilder.Append "UPDATE org_cost SET "
objBuilder.Append "	cost_amt_"&cost_month&" = '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND bonbu ='"&bonbu&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'상주비용 마감 취소
objBuilder.Append "UPDATE cost_end SET "
objBuilder.Append "	end_yn = 'C', batch_yn = 'N', bonbu_yn = 'N', "
objBuilder.Append "	mod_id = '"&user_id&"', mod_name = '"&user_name&"', mod_date = NOW() "
objBuilder.Append "WHERE end_month = '"&end_month&"' "
objBuilder.Append "	AND saupbu = '상주비용' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "UPDATE company_cost SET "
objBuilder.Append "	cost_amt_"&cost_month&"= '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "UPDATE company_profit_loss SET "
objBuilder.Append "	cost_amt_"&cost_month&"= '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "UPDATE saupbu_profit_loss SET "
objBuilder.Append "	cost_amt_"&cost_month&"= '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 공통비 배분 취소
objBuilder.Append "UPDATE cost_end SET "
objBuilder.Append "	end_yn = 'C', batch_yn = 'N', bonbu_yn = 'N', "
objBuilder.Append "	mod_id = '"&user_id&"', mod_name = '"&user_name&"', mod_date = NOW() "
objBuilder.Append "WHERE end_month = '"&end_month&"' "
objBuilder.Append "	AND saupbu = '공통비/직접비배분' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "DELETE FROM company_as "
objBuilder.Append "WHERE as_month = '"&end_month&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' AS 표준단가
objBuilder.Append "DELETE FROM company_asunit "
objBuilder.Append "WHERE as_month = '"&end_month&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "DELETE FROM management_cost "
objBuilder.Append "WHERE cost_month = '"&end_month&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "처리중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "마감이 취소되었습니다."
End If

Response.Write "<script language=javascript>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('/cost_end/cost_end_mg.asp');"
Response.Write "</script>"
Response.End

DBConn.Close() : Set DBConn = Nothing
%>


