<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	account_group = request.form("account_group")
	account_seq = request.form("account_seq")
	account_name = request.form("account_name")
	item_seq = request.form("item_seq")
	account_item = request.form("account_item")
	cost_yn = request.form("cost_yn")

	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "Update account_item set account_item='"&account_item&"',cost_yn='"&cost_yn&"',mod_user='"&user_name&"',mod_date=now() where account_group = '"+account_group+"' and account_seq ='"+account_seq+"' and item_seq ='"+item_seq+"'"
		dbconn.execute(sql)
	  else
		sql="select max(item_seq) as max_seq from account_item where account_group = '"+account_group+"' and account_seq ='"+account_seq+"'"
		Set rs=DbConn.Execute(Sql)
		if rs("max_seq") = "" or isnull(rs("max_seq")) then
			max_seq = "01"
		  else
		  	max_seq = cint(rs("max_seq")) + 1
			if max_seq < 10 then
				max_seq = "0" + cstr(max_seq)
			  else
			  	max_seq = cstr(max_seq)
			end if
		end if
		
		sql="insert into account_item (account_group,account_seq,item_seq,account_name,account_item,cost_yn,reg_user,reg_date) "
		sql=sql + "values ('"&account_group&"','"&account_seq&"','"&max_seq&"','"&account_name&"','"&account_item&"','"&cost_yn&"','"&user_name&"',now())"
		dbconn.execute(sql)
	end if	

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('등록 완료 되었습니다....');"		
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
