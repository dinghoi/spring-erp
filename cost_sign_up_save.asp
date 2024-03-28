<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
<%
'	on Error resume next

	sign_month = request.form("sign_month")
	team_emp_no = "mst01"
	saupbu_emp_no = "mst02"
	bonbu_emp_no = "mst03"	
	ceo_emp_no = "mst04"	
	sign_memo = request.form("sign_memo")
	sign_pro = request.form("sign_pro")
	sign_id = request.form("sign_id")
	sign_head = request.form("sign_head")
	from_date = request.form("from_date")
	to_date = request.form("to_date")

	if position = "팀장" then
		sign_pro = "T"
		team_sign = "I"
		saupbu_sign = "I"
		bonbu_sign = "I"
		ceo_sign = "I"		
	end if
	if position = "사업부장" then
		sign_pro = "S"
		team_sign = "N"
		saupbu_sign = "I"
		bonbu_sign = "I"
		ceo_sign = "I"
	end if
	if position = "본부장" then
		sign_pro = "B"
		team_sign = "N"
		saupbu_sign = "N"
		bonbu_sign = "I"
		ceo_sign = "I"
	end if
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sign_date = mid(now(),1,10)
	sql="select max(sign_seq) as max_seq from sign_process where sign_date='" + sign_date + "'"
	set rs=dbconn.execute(sql)
		
	if	isnull(rs("max_seq"))  then
		sign_seq = "001"
	  else
		max_seq = "00" + cstr((int(rs("max_seq")) + 1))
		sign_seq = right(max_seq,3)
	end if

	if position = "팀장" then
		sql = "insert into sign_process (sign_date,sign_seq,sign_month,sign_id,bonbu,saupbu,team,sign_pro,team_emp_no,team_sign,team_date"& _
		",saupbu_emp_no,saupbu_sign,bonbu_emp_no,bonbu_sign,ceo_emp_no,ceo_sign,sign_memo,reg_id,reg_user,reg_date) values "& _
		" ('"&sign_date&"','"&sign_seq&"','"&sign_month&"','"&sign_id&"','"&bonbu&"','"&saupbu&"','"&team&"','"&sign_pro&"','"&team_emp_no& _
		"','"&team_sign&"',now(),'"&saupbu_emp_no&"','"&saupbu_sign&"','"&bonbu_emp_no&"','"&bonbu_sign&"','"&ceo_emp_no&"','"&ceo_sign& _
		"','"&sign_memo&"','"&user_id&"','"&user_name&"',now())"
	end if
	if position = "사업부장" then
		sql = "insert into sign_process (sign_date,sign_seq,sign_month,sign_id,bonbu,saupbu,sign_pro,team_emp_no,team_sign"& _
		",saupbu_emp_no,saupbu_sign,saupbu_date,bonbu_emp_no,bonbu_sign,ceo_emp_no,ceo_sign,sign_memo,reg_id,reg_user,reg_date) values "& _
		" ('"&sign_date&"','"&sign_seq&"','"&sign_month&"','"&sign_id&"','"&bonbu&"','"&saupbu&"','"&sign_pro&"','"&team_emp_no& _
		"','"&team_sign&"','"&saupbu_emp_no&"','"&saupbu_sign&"',now(),'"&bonbu_emp_no&"','"&bonbu_sign&"','"&ceo_emp_no&"','"&ceo_sign& _
		"','"&sign_memo&"','"&user_id&"','"&user_name&"',now())"
	end if
	if position = "본부장" then
		sql = "insert into sign_process (sign_date,sign_seq,sign_month,sign_id,bonbu,sign_pro,team_emp_no,team_sign"& _
		",saupbu_emp_no,saupbu_sign,bonbu_emp_no,bonbu_sign,bonbu_date,ceo_emp_no,ceo_sign,sign_memo,reg_id,reg_user,reg_date) values "& _
		" ('"&sign_date&"','"&sign_seq&"','"&sign_month&"','"&sign_id&"','"&bonbu&"','"&sign_pro&"','"&team_emp_no& _
		"','"&team_sign&"','"&saupbu_emp_no&"','"&saupbu_sign&"','"&bonbu_emp_no&"','"&bonbu_sign&"',now(),'"&ceo_emp_no&"','"&ceo_sign& _
		"','"&sign_memo&"','"&user_id&"','"&user_name&"',now())"
	end if
	
	dbconn.execute(sql)
	paper_no = sign_date + "-" + sign_seq

	sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
	sql = sql +	" ('"&user_id&"','"&user_name&"','"&user_id&"','"&paper_no&"','"&sign_head&"','N','N',now())"
	dbconn.execute(sql)

	if position = "팀장" then
		sql = "update general_cost set end_yn='I',paper_no='"&paper_no&"' where (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and bonbu = '"&bonbu&"' and saupbu = '"&saupbu&"' and team = '"&team&"'"
	end if
	if position = "사업부장" or position = "본부장" then
		sql = "update general_cost set end_yn='I',paper_no='"&paper_no&"' where (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and reg_id = '"&user_id&"'"
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
