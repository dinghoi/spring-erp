<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

emp_user = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

	u_type = request.form("u_type")
	pe_seq = request.form("pe_seq")
	
	car_no = request.form("car_no")
	car_name = request.form("car_name")
	car_year = request.form("car_year")
	car_reg_date = request.form("car_reg_date")
	owner_emp_name = request.form("owner_emp_name")
	owner_emp_no = request.form("owner_emp_no")
	car_use_dept = request.form("car_use_dept")
    car_owner = request.form("car_owner")
	
	pe_car_no = request.form("car_no")
	pe_date = request.form("pe_date")

    pe_comment = request.form("pe_comment")
    pe_place = request.form("pe_place")
    pe_amount = int(request.form("pe_amount"))
    pe_in_date = request.form("pe_in_date")
	pe_in_amt = int(request.form("pe_in_amt"))
    pe_default = request.form("pe_default")
    pe_notice_date = request.form("pe_notice_date")
    pe_notice = request.form("pe_notice")
    pe_bigo = request.form("pe_bigo")
	
	if pe_notice_date = "" or isnull(pe_notice_date) then
	   pe_notice_date = "1900-01-01"
	end if
	if pe_in_date = "" or isnull(pe_in_date) then
	   pe_in_date = "1900-01-01"
	end if
	
	set dbconn = server.CreateObject("adodb.connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
    Set Rs_pe = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	Set Rs_trans = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

    if u_type <> "U" then
		sql="select max(pe_seq) as max_seq from car_penalty where pe_car_no='" + pe_car_no + "' and pe_date='" + pe_date + "'"
		set Rs_pe=dbconn.execute(sql)
		if	isnull(Rs_pe("max_seq"))  then
			pe_seq = "001"
		  else
			max_seq = "00" + cstr((int(Rs_pe("max_seq")) + 1))
			pe_seq = right(max_seq,3)
		end if	 
		Rs_pe.close()
	end if
		
	dbconn.BeginTrans

	if	u_type = "U" then
		sql = "Update car_penalty set pe_comment='"&pe_comment&"',pe_place='"&pe_place&"',pe_amount='"&pe_amount&"',pe_in_date='"&pe_in_date&"',pe_in_amt='"&pe_in_amt&"',pe_default='"&pe_default&"',pe_notice_date='"&pe_notice_date&"',pe_notice='"&pe_notice&"',pe_bigo='"&pe_bigo&"' where pe_car_no = '"&pe_car_no&"' and pe_date = '"&pe_date&"' and pe_seq = '"&pe_seq&"'"
		
		dbconn.execute(sql)
				
	  else
  
		sql="insert into car_penalty (pe_car_no,pe_date,pe_seq,pe_car_name,pe_car_owner,pe_use_dept,pe_owner_emp_no,pe_owner_emp_name,pe_comment,pe_place,pe_amount,pe_in_date,pe_in_amt,pe_default,pe_notice_date,pe_notice,pe_bigo,pe_reg_date,pe_reg_user) values ('"&pe_car_no&"','"&pe_date&"','"&pe_seq&"','"&car_name&"','"&car_owner&"','"&car_use_dept&"','"&owner_emp_no&"','"&owner_emp_name&"','"&pe_comment&"','"&pe_place&"','"&pe_amount&"','"&pe_in_date&"','"&pe_in_amt&"','"&pe_default&"','"&pe_notice_date&"','"&pe_notice&"','"&pe_bigo&"',now(),'"&emp_user&"')"
		dbconn.execute(sql)
	
	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = sms_msg + "자장중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = sms_msg + "저장되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	

%>
