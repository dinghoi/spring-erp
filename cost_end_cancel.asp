<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	org_company=Request("org_company")
	saupbu=Request("saupbu")
	end_month=Request("end_month")
	end_yn=Request("end_yn")

	cost_year = mid(end_month,1,4)
	cost_month = mid(end_month,5)
	
	from_date = mid(end_month,1,4) + "-" + mid(end_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))
	
	response.write"<script language=javascript>"
	response.write"alert('마감 취소중!!!');"
	response.write"</script>"

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	dbconn.BeginTrans

'야특근 마감
    sql = "select * from overtime where work_date >= '"&from_date&"' and work_date <= '"&to_date&"' and saupbu ='"&saupbu&"'"
	Rs.Open Sql, Dbconn, 1

	do until rs.eof
		sql = "Update overtime set end_yn='C' where work_date = '"&rs("work_date")&"' and mg_ce_id = '"&rs("mg_ce_id")&"'"
		dbconn.execute(sql)
		rs.movenext()
	loop
	rs.close()

'일반비용	
    sql = "select * from general_cost where (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and saupbu ='"&saupbu&"'"
	Rs.Open Sql, Dbconn, 1

	do until rs.eof
		sql = "Update general_cost set end_yn='C' where slip_date = '"&rs("slip_date")&"' and slip_seq = '"&rs("slip_seq")&"'"
		dbconn.execute(sql)
		rs.movenext()
	loop
	rs.close()

'교통비
	sql = "select * from transit_cost where (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and saupbu ='"&saupbu&"'"
    Rs.Open Sql, Dbconn, 1

	do until rs.eof
		sql = "Update transit_cost set end_yn='C' where run_date = '"&rs("run_date")&"' and mg_ce_id = '"&rs("mg_ce_id")&"' and run_seq ='"&rs("run_seq")&"'"
		dbconn.execute(sql)
		rs.movenext()
	loop
	rs.close()

	sql = "Update cost_end set end_yn='C',batch_yn='N',bonbu_yn='N',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where end_month = '"&end_month& _
    "' and saupbu = '"&saupbu&"'"
	dbconn.execute(sql)
	
	sql = "update org_cost set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and saupbu ='"&saupbu&"'"
    dbconn.execute(sql)

'	sql = "Update cost_end set end_yn='C',batch_yn='N',bonbu_yn='N',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where end_month = '"&end_month&"' and saupbu = '상주비용'"
'	dbconn.execute(sql)

'	sql = "update company_cost set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
'	dbconn.execute(sql)

' 상주비용 취소
	sql = "Update cost_end set end_yn='C',batch_yn='N',bonbu_yn='N',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where end_month = '"&end_month&"' and saupbu = '상주비용'"
    dbconn.execute(sql)

    sql = "update company_cost set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
	dbconn.execute(sql)
    sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
	dbconn.execute(sql)
    sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
	dbconn.execute(sql)

' 공통비 배분 취소
    sql = "Update cost_end set end_yn='C',batch_yn='N',bonbu_yn='N',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where end_month = '"&end_month&"' and saupbu = '공통비/직접비배분'"
	dbconn.execute(sql)

    sql = "delete from company_as where as_month ='"&end_month&"'"
	dbconn.execute(sql)
	sql = "delete from company_asunit where as_month ='"&end_month&"'" ' AS 표준단가
	dbconn.execute(sql)
    sql = "delete from management_cost where cost_month ='"&end_month&"'"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "처리중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "마감이 취소되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('cost_end_mg.asp');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>


