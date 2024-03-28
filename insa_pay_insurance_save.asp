<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	insu_id = request.form("insu_id")
	insu_class = request.form("insu_class")
	insu_id_name = request.form("insu_id_name")

  insu_yyyy = request.form("insu_yyyy")
	'if	u_type = "U" then
	'       insu_yyyy = request.form("insu_yyyy")
	'else
	'       from_date = request.form("from_date")
	'       insu_yyyy = mid(cstr(from_date),1,4)
  'end if
	
	from_amt = int(request.form("from_amt"))
	to_amt = int(request.form("to_amt"))
	st_amt = int(request.form("st_amt"))
	
	emp_rate = request.form("emp_rate")
	com_rate = request.form("com_rate")
	tot_rate = request.form("hap_rate")
	'hap_rate = emp_rate + com_rate
	insu_comment = request.form("insu_comment")
	
	'start_time = cstr(start_hh) + cstr(start_mm)
	
	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "Update pay_insurance set from_amt='"&from_amt&"',to_amt ='"&to_amt&"',st_amt ='"&st_amt&"',hap_rate='"&tot_rate&"',emp_rate='"&emp_rate&"',com_rate='"&com_rate&"',insu_comment='"&insu_comment&"',mod_user='"&emp_user&"',mod_date=now() where insu_yyyy = '"&insu_yyyy&"' and insu_id = '"&insu_id&"' and insu_class = '"&insu_class&"'"
		dbconn.execute(sql)
		
	  else
		sql="insert into pay_insurance (insu_yyyy,insu_id,insu_class,insu_id_name,from_amt,to_amt,st_amt,hap_rate,emp_rate,com_rate,insu_comment,reg_user,reg_date) values ('"&insu_yyyy&"','"&insu_id&"','"&insu_class&"','"&insu_id_name&"','"&from_amt&"','"&to_amt&"','"&st_amt&"','"&tot_rate&"','"&emp_rate&"','"&com_rate&"','"&insu_comment&"','"&emp_user&"',now())"
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
