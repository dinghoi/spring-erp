<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	
	fm_id = request.form("fm_id")
	fm_type = request.form("fm_type")
	fm_sawo_pay = request.form("fm_sawo_pay")
    fm_company_pay = request.form("fm_company_pay")
    fm_holiday1 = request.form("fm_holiday1")
    fm_holiday2 = request.form("fm_holiday2")
    fm_wreath_yn = request.form("fm_wreath_yn")
    fm_flowers_yn = request.form("fm_flowers_yn")
    fm_comment = request.form("fm_comment")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

emp_user = request.cookies("nkpmg_user")("coo_user_name")

	if	u_type = "U" then
		sql = "update emp_family_event set fm_sawo_pay='"&fm_sawo_pay&"',fm_company_pay='"&fm_company_pay&"',fm_holiday1='"&fm_holiday1&"',fm_holiday2='"&fm_holiday2&"',fm_wreath_yn='"&fm_wreath_yn&"',fm_flowers_yn='"&fm_flowers_yn&"',fm_comment='"&fm_comment&"',fm_mod_date=now(),fm_mod_user='"&emp_user&"' where fm_id ='"&fm_id&"' and fm_type = '"&fm_type&"'"
		
		'response.write sql
		dbconn.execute(sql)	  
	  else
		sql = "insert into emp_family_event (fm_id,fm_type,fm_sawo_pay,fm_company_pay,fm_holiday1,fm_holiday2,fm_wreath_yn,fm_flowers_yn,fm_comment,fm_reg_date,fm_reg_user) values "
		sql = sql +	" ('"&fm_id&"','"&fm_type&"','"&fm_sawo_pay&"','"&fm_company_pay&"','"&fm_holiday1&"','"&fm_holiday2&"','"&fm_wreath_yn&"','"&fm_flowers_yn&"','"&fm_comment&"',now(),'"&emp_user&"')"
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
