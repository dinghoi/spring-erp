<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	from_month=Request("from_month")
	to_month=Request("to_month")

	from_date = mid(from_month,1,4) + "-" + mid(from_month,5,2)
	to_date = mid(to_month,1,4) + "-" + mid(to_month,5,2)
	
	response.write"<script language=javascript>"
	response.write"alert('���� �����!!!');"
	response.write"</script>"

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	dbconn.BeginTrans

'��Ư�� ����
	sql = "Update overtime set end_yn='N' where substring(work_date,1,7) >= '"&from_date&"' and substring(work_date,1,7) <= '"&to_date&"'"
	dbconn.execute(sql)

'�Ϲݺ��	
	sql = "Update general_cost set end_yn='N' where substring(slip_date,1,7) >= '"&from_date&"' and substring(slip_date,1,7) <= '"&to_date&"'"
	dbconn.execute(sql)

'�����
	sql = "Update transit_cost set end_yn='N' where substring(run_date,1,7) >= '"&from_date&"' and substring(run_date,1,7) <= '"&to_date&"'"
	dbconn.execute(sql)

' ���� ������ ����
	sql = "delete from cost_end where end_month >= '"&from_month&"' and end_month <= '"&to_month&"'"
	dbconn.execute(sql)

' ������ ��� CLEAR	
	for yyyymm = from_month to to_month
		cost_year = mid(yyyymm,1,4)
		cost_month = cstr(mid(yyyymm,5))

		sql = "update org_cost set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
		dbconn.execute(sql)

		sql = "update company_cost set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
		dbconn.execute(sql)

		sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
		dbconn.execute(sql)

		sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"'"
		dbconn.execute(sql)
	next

' ����� ��� ���
	sql = "delete from company_as where as_month >= '"&from_month&"' and as_month <= '"&to_month&"'"
	dbconn.execute(sql)
	sql = "delete from company_asunit where as_month >= '"&from_month&"' and as_month ='"&to_month&"'" ' AS ǥ�شܰ�
	dbconn.execute(sql)
	sql = "delete from management_cost where cost_month >= '"&from_month&"' and cost_month <= '"&to_month&"'"
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "ó���� Error�� �߻��Ͽ����ϴ�...."
	else    
		dbconn.CommitTrans
		end_msg = "������ ��ҵǾ����ϴ�...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('cost_end_mg.asp');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>


