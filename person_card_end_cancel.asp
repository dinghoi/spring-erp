<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	slip_month=Request("slip_month")
	emp_no=Request("emp_no")
	ck_sw = "y"
		
	from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))

	response.write"<script language=javascript>"
	response.write"    var yes_no = confirm('���� ���ó���Ͻðڽ��ϱ�?');"
	response.write"    if(yes_no==false){"
	response.write"        alert('��ҵǾ����ϴ�');"
	response.write"        history.back();"
	response.write"    }"
	response.write"</script>"

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	dbconn.BeginTrans

'��Ư�� ����
	sql = "select * from card_slip where slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"' and emp_no ='"&emp_no&"'"
	'response.write(sql)
	Rs.Open Sql, Dbconn, 1

	do until rs.eof
		sql = "Update card_slip set person_end='N' where approve_no = '"&rs("approve_no")&"'"
        dbconn.execute(sql)
        
		rs.movenext()
	loop
	rs.close()
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "ó���� Error�� �߻��Ͽ����ϴ�...."
	else    
		dbconn.CommitTrans
		end_msg = "�������ó�� �Ǿ����ϴ�...."
	end if

	url = "person_card_mg.asp?slip_month="&slip_month&"&ck_sw="&ck_sw

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>


