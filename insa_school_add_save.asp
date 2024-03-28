<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
	on Error resume next

	u_type = request.form("u_type")
	
	sch_seq = request.form("sch_seq")
	sch_empno = request.form("sch_empno")
	view_condi = request.form("view_condi")
	if view_condi = "1" then 
	         sch_school_name = request.form("sch_high_name")
	   else 
	         sch_school_name = request.form("sch_school_name")
	end if
	
	sch_start_date = request.form("sch_start_date")
    sch_end_date = request.form("sch_end_date")
    sch_dept = request.form("sch_dept")
    sch_major = request.form("sch_major")
    sch_sub_major = request.form("sch_sub_major")
    sch_degree = request.form("sch_degree")
	sch_finish = request.form("sch_finish")
	sch_comment = view_condi
    'sch_comment = request.form("sch_comment")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update emp_school set sch_start_date='"&sch_start_date&"',sch_end_date='"&sch_end_date&"',sch_school_name='"&sch_school_name&"',sch_dept='"&sch_dept&"',sch_major='"&sch_major&"',sch_sub_major='"&sch_sub_major&"',sch_degree='"&sch_degree&"',sch_finish='"&sch_finish&"',sch_comment='"&sch_comment&"',sch_mod_date= now(),sch_mod_user='"&emp_user&"' where sch_empno ='"&sch_empno&"' and sch_seq = '"&sch_seq&"'"
		
		'response.write sql
		
		dbconn.execute(sql)	  
	  else
		sql="select max(sch_seq) as max_seq from emp_school where sch_empno='" + sch_empno + "'"
		set rs=dbconn.execute(sql)
		
		if	isnull(rs("max_seq"))  then
			sch_seq = "001"
		  else
			max_seq = "00" + cstr((int(rs("max_seq")) + 1))
			sch_seq = right(max_seq,3)
		end if

		sql = "insert into emp_school (sch_empno,sch_seq,sch_start_date,sch_end_date,sch_school_name,sch_dept,sch_major,sch_sub_major,sch_degree,sch_finish,sch_comment,sch_reg_date,sch_reg_user) values "
		sql = sql +	" ('"&sch_empno&"','"&sch_seq&"','"&sch_start_date&"','"&sch_end_date&"','"&sch_school_name&"','"&sch_dept&"','"&sch_major&"','"&sch_sub_major&"','"&sch_degree&"','"&sch_finish&"','"&sch_comment&"',now(),'"&emp_user&"')"
		
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
