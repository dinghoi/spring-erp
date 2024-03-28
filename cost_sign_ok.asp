<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
<%
'	on Error resume next

	sign_date=request.form("sign_date")
	sign_seq=request.form("sign_seq")
	sign_month=request.form("sign_month")
	msg_seq=int(request.form("msg_seq"))
	sign_yn=request.form("sign_yn")
	sign_memo=request.form("sign_memo")
	title_line=request.form("title_line") + " 반려"
	paper_no = cstr(sign_date) + "-" + cstr(sign_seq)

	from_date = cstr(mid(sign_month,1,4)) + "-" + cstr(mid(sign_month,5,2)) + "-" + "01"
	to_date = cstr(mid(sign_month,1,4)) + "-" + cstr(mid(sign_month,5,2)) + "-" + "31"

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sql="select * from sign_msg where msg_seq="&msg_seq
	set rs_msg=dbconn.execute(sql)

	sql="select * from sign_process where sign_date='"&sign_date&"' and sign_seq ='"&sign_seq&"'"
	set rs=dbconn.execute(sql)

	if rs("sign_pro") = "T" and rs("team_sign") = "I" then
		if sign_yn = "C" then		
			sql = "update sign_process set sign_pro='C',team_sign='C',team_date=now(),sign_memo='"&sign_memo&"' where sign_date='"&sign_date&"' and sign_seq='"&sign_seq&"'"
			dbconn.execute(sql)

			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&user_id&"','"&user_name&"','"&rs("reg_id")&"','"&paper_no&"','"&title_line&"','N','C',now())"
			dbconn.execute(sql)

		  else	
			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&rs_msg("send_id")&"','"&rs_msg("send_name")&"','"&rs("saupbu_emp_no")&"','"&rs_msg("paper_no")&"','"&rs_msg("sign_head")&"','N','N',now())"
			dbconn.execute(sql)
			sql = "update sign_process set sign_pro='S',team_sign='E',team_date=now(),sign_memo='"&sign_memo&"' where sign_date='"&sign_date&"' and sign_seq='"&sign_seq&"'"
			dbconn.execute(sql)
		end if
	end if

	if rs("sign_pro") = "S" and rs("saupbu_sign") = "I" then
		if sign_yn = "C" then		
			sql = "update sign_process set sign_pro='C',saupbu_sign='C',saupbu_date=now(),sign_memo='"&sign_memo&"' where sign_date='"&sign_date&"' and sign_seq='"&sign_seq&"'"
			dbconn.execute(sql)

			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&user_id&"','"&user_name&"','"&rs("reg_id")&"','"&paper_no&"','"&title_line&"','N','C',now())"
			dbconn.execute(sql)
			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&user_id&"','"&user_name&"','"&rs("team_emp_no")&"','"&paper_no&"','"&title_line&"','N','C',now())"
			dbconn.execute(sql)
		  else	
			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&rs_msg("send_id")&"','"&rs_msg("send_name")&"','"&rs("bonbu_emp_no")&"','"&rs_msg("paper_no")&"','"&rs_msg("sign_head")&"','N','N',now())"
			dbconn.execute(sql)
			sql = "update sign_process set sign_pro='B',saupbu_sign='E',saupbu_date=now(),sign_memo='"&sign_memo&"' where sign_date='"&sign_date&"' and sign_seq='"&sign_seq&"'"
			dbconn.execute(sql)
		end if
	end if
	
	if rs("sign_pro") = "B" and rs("bonbu_sign") = "I" then
		if sign_yn = "C" then		
			sql = "update sign_process set sign_pro='C',bonbu_sign='C',bonbu_date=now(),sign_memo='"&sign_memo&"' where sign_date='"&sign_date&"' and sign_seq='"&sign_seq&"'"
			dbconn.execute(sql)

			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&user_id&"','"&user_name&"','"&rs("reg_id")&"','"&paper_no&"','"&title_line&"','N','C',now())"
			dbconn.execute(sql)
			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&user_id&"','"&user_name&"','"&rs("team_emp_no")&"','"&paper_no&"','"&title_line&"','N','C',now())"
			dbconn.execute(sql)
			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&user_id&"','"&user_name&"','"&rs("saupbu_emp_no")&"','"&paper_no&"','"&title_line&"','N','C',now())"
			dbconn.execute(sql)
		  else	
			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&rs_msg("send_id")&"','"&rs_msg("send_name")&"','"&rs("ceo_emp_no")&"','"&rs_msg("paper_no")&"','"&rs_msg("sign_head")&"','N','N',now())"
			dbconn.execute(sql)
			sql = "update sign_process set sign_pro='O',bonbu_sign='E',bonbu_date=now(),sign_memo='"&sign_memo&"' where sign_date='"&sign_date&"' and sign_seq='"&sign_seq&"'"
			dbconn.execute(sql)
		end if
	end if
	
	if rs("sign_pro") = "O" and rs("ceo_sign") = "I" then
		if sign_yn = "C" then		
			response.write("aaaa")
			sql = "update sign_process set sign_pro='C',ceo_sign='C',ceo_date=now(),sign_memo='"&sign_memo&"' where sign_date='"&sign_date&"' and sign_seq='"&sign_seq&"'"
			dbconn.execute(sql)

			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&user_id&"','"&user_name&"','"&rs("reg_id")&"','"&paper_no&"','"&title_line&"','N','C',now())"
			dbconn.execute(sql)
			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&user_id&"','"&user_name&"','"&rs("team_emp_no")&"','"&paper_no&"','"&title_line&"','N','C',now())"
			dbconn.execute(sql)
			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&user_id&"','"&user_name&"','"&rs("saupbu_emp_no")&"','"&paper_no&"','"&title_line&"','N','C',now())"
			dbconn.execute(sql)
			sql = "insert into sign_msg (send_id,send_name,recv_id,paper_no,sign_head,read_yn,sign_yn,reg_date) values "
			sql = sql +	" ('"&user_id&"','"&user_name&"','"&rs("bonbu_emp_no")&"','"&paper_no&"','"&title_line&"','N','C',now())"
			dbconn.execute(sql)
		  else	
			sql = "update sign_process set sign_pro='E',ceo_sign='E',ceo_date=now(),sign_memo='"&sign_memo&"' where sign_date='"&sign_date&"' and sign_seq='"&sign_seq&"'"
			dbconn.execute(sql)
		end if
	end if

	sql = "update sign_msg set sign_yn='Y' where msg_seq="&int(msg_seq)
	dbconn.execute(sql)	  

	if rs("sign_pro") = "O" and sign_yn = "Y" then		
		sql = "update general_cost set end_yn='Y' where paper_no ='"&paper_no&"'"
		dbconn.execute(sql)	  
	end if
	if sign_yn = "C" then
		sql = "update general_cost set end_yn='C' where paper_no ='"&paper_no&"'"
		dbconn.execute(sql)	  
	end if

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "결재중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "결재되었습니다...."
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
