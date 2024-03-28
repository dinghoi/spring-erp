<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	collect_id = request.form("collect_id")
	approve_no = request.form("approve_no")
	slip_no = request.form("slip_no")
	collect_date = request.form("collect_date")
	bill_collect = request.form("bill_collect")
	bill_date = request.form("bill_date")
	unpaid_due_date = request.form("unpaid_due_date")
	unpaid_memo = request.form("unpaid_memo")
	unpaid_memo = Replace(unpaid_memo,"'","&quot;")
	collect_amt = int(request.form("collect_amt"))
	collect_tot_amt = int(request.form("collect_tot_amt"))
	sales_amt = int(request.form("sales_amt"))
	collect_due_date = request.form("collect_due_date")
	change_memo = request.form("change_memo")
	change_memo1 = request.form("change_memo1")
	end_date = request.form("end_date")
	curr_date = mid(now(),1,10)
	
	if collect_amt = "" or isnull(collect_amt) then
		collect_amt = 0
	end if
	collect_tot_amt = collect_tot_amt + collect_amt

	if bill_date = "" or isnull(bill_date) then
		bill_date = "0000-00-00"
	end if
	if unpaid_due_date = "" or isnull(unpaid_due_date) then
		unpaid_due_date = "0000-00-00"
	end if
	 
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sql="select max(collect_seq) as max_seq from sales_collect where approve_no='"&approve_no&"'" 
	set rs=dbconn.execute(sql)
		
	if	isnull(rs("max_seq"))  then
		collect_seq = "01"
	  else
		max_seq = "0" + cstr((int(rs("max_seq")) + 1))
		collect_seq = right(max_seq,2)
	end if

	if collect_id = "1" then
		sql = "insert into sales_collect (approve_no,collect_seq,slip_no,collect_date,collect_id,bill_collect,collect_amt,bill_date,collect_due_date,unpaid_memo,reg_emp_no,reg_name,reg_date) values ('"&approve_no&"','"&collect_seq&"','"&slip_no&"','"&collect_date&"','"&collect_id&"','"&bill_collect&"',"&collect_amt&",'"&bill_date&"','"&collect_due_date&"','','"&user_id&"','"&user_name&"',now())"
	end if
	if collect_id = "2" or collect_id = "3" then
		sql = "insert into sales_collect (approve_no,collect_seq,slip_no,collect_date,collect_id,bill_collect,collect_amt,collect_due_date,change_memo,unpaid_memo,unpaid_due_date,reg_emp_no,reg_name,reg_date) values ('"&approve_no&"','"&collect_seq&"','"&slip_no&"','"&curr_date&"','"&collect_id&"','',0,'"&collect_due_date&"','"&change_memo&"','"&unpaid_memo&"','"&unpaid_due_date&"','"&user_id&"','"&user_name&"',now())"
	end if
	if collect_id = "4" then
		sql = "insert into sales_collect (approve_no,collect_seq,slip_no,collect_date,collect_id,bill_collect,collect_amt,change_memo,unpaid_memo,reg_emp_no,reg_name,reg_date) values ('"&approve_no&"','"&collect_seq&"','"&slip_no&"','"&end_date&"','"&collect_id&"','',0,'"&change_memo1&"','','"&user_id&"','"&user_name&"',now())"
	end if

	dbconn.execute(sql)

	if collect_id = "1" then
		if sales_amt = collect_tot_amt then
			slip_stat = "1"
		  elseif sales_amt < collect_tot_amt then	
		    slip_stat = "2"
		  else
		  	slip_stat = "0"
		end if	
		sql = "update saupbu_sales set collect_tot_amt="&collect_tot_amt&", slip_stat='"&slip_stat&"', mod_id='"&user_id&"', mod_name='"&user_name&"', mod_date=now() where approve_no='"&approve_no&"'" 
	end if

	if collect_id = "2" then
		sql = "update saupbu_sales set change_memo='"&change_memo&"', unpaid_memo='"&unpaid_memo&"', unpaid_due_date='"&unpaid_due_date&"', slip_stat='0', mod_id='"&user_id&"', mod_name='"&user_name&"', mod_date=now() where approve_no='"&approve_no&"'" 
	end if

	if collect_id = "3" then
		sql = "update saupbu_sales set collect_due_date='"&unpaid_due_date&"', change_memo='"&change_memo&"', unpaid_memo='"&unpaid_memo&"', unpaid_due_date='0000-00-00', slip_stat='0', mod_id='"&user_id&"', mod_name='"&user_name&"', mod_date=now() where approve_no='"&approve_no&"'" 
	end if

	if collect_id = "4" then
		sql = "update saupbu_sales set change_memo='"&change_memo1&"', slip_stat='1', mod_id='"&user_id&"', mod_name='"&user_name&"', mod_date=now() where approve_no='"&approve_no&"'" 
	end if
	
	dbconn.execute(sql)	  

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
