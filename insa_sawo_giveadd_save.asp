<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	u_type = request.form("u_type")
	ask_seq = request.form("ask_seq")
	
	give_empno = request.form("give_empno")
    give_seq = request.form("give_seq")
    give_date = request.form("give_date")
	give_ask_process = request.form("give_ask_process")
    give_emp_name = request.form("give_emp_name")
    give_company = request.form("give_company")
    give_org = request.form("give_org")
    give_org_name = request.form("give_org_name")
    give_pay = int(request.form("give_pay"))
	give_comment = request.form("give_comment")
	give_id = request.form("give_id")
    give_type = request.form("give_type")
    give_sawo_date = request.form("give_sawo_date")
    give_sawo_place = request.form("give_sawo_place")
    give_sawo_comm = request.form("give_sawo_comm")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs_mem = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

' 경조회회원마스터 update를 위한...
if give_ask_process = "2" then
    Sql="select * from emp_sawo_mem where sawo_empno = '"&give_empno&"'"
    Rs_mem.Open Sql, Dbconn, 1
 
    sawo_give_cnt = 0
	sawo_give_pay = 0
    if not Rs_mem.eof then
       sawo_give_cnt = Rs_mem("sawo_give_count")
       sawo_give_pay = Rs_mem("sawo_give_pay")
    end if
    Rs_mem.Close()

    sawo_give_cnt = sawo_give_cnt + 1
    sawo_give_pay = sawo_give_pay + give_pay
end if

	dbconn.BeginTrans


	if	u_type = "U" then
		sql = "update emp_sawo_give set give_pay='"&give_pay&"',give_comment='"&give_comment&"',give_id='"&give_id&"',give_type='"&give_type&"',give_sawo_date='"&give_sawo_date&"',give_sawo_place='"&give_sawo_place&"',give_sawo_comm='"&give_sawo_comm&"',give_mod_date=now(),give_mod_user='"&emp_user&"' where give_empno ='"&give_empno&"' and give_seq = '"&give_seq&"' and give_date = '"&give_date&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql = "insert into emp_sawo_give (give_empno,give_seq,give_date,give_ask_process,give_emp_name,give_company,give_org,give_org_name,give_id,give_type,give_pay,give_sawo_date,give_sawo_place,give_sawo_comm,give_comment,give_reg_date,give_reg_user) values "
		sql = sql +	" ('"&give_empno&"','"&give_seq&"','"&give_date&"','"&give_ask_process&"','"&give_emp_name&"','"&give_company&"','"&give_org&"','"&give_org_name&"','"&give_id&"','"&give_type&"','"&give_pay&"','"&give_sawo_date&"','"&give_sawo_place&"','"&give_sawo_comm&"','"&give_comment&"',now(),'"&emp_user&"')"
		dbconn.execute(sql)
	end if

if give_ask_process = "2" then	
	    sql = "update emp_sawo_mem set sawo_give_count='"&sawo_give_cnt&"',sawo_give_pay='"&sawo_give_pay&"' where sawo_empno ='"&give_empno&"'"
		dbconn.execute(sql)	  
		
		sql = "update emp_sawo_ask set ask_process='1' where ask_empno ='"&give_empno&"' and ask_seq ='"&ask_seq&"' and ask_date ='"&give_sawo_date&"'"
		'response.write sql
		
		dbconn.execute(sql)	  
end if

if give_ask_process = "1" then	
		sql = "update emp_sawo_ask set ask_company_process='1' where ask_empno ='"&give_empno&"' and ask_seq ='"&ask_seq&"' and ask_date ='"&give_sawo_date&"'"
		'response.write sql
		
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
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

	
%>
