<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next
dim abc,filenm

Set abc = Server.CreateObject("ABCUpload4.XForm")
abc.AbsolutePath = True
abc.Overwrite = true
abc.MaxUploadSize = 1024*1024*50

    emp_user = request.cookies("nkpmg_user")("coo_user_name")
    emp_no = request.cookies("nkpmg_user")("coo_user_id")

    pay_company = abc("pay_company1")
    pay_month   = abc("pay_month1")
	
'	response.write(pay_company)
'	response.write(pay_month)
'	response.End

	set dbconn = server.CreateObject("adodb.connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set Rs_emp = Server.CreateObject("ADODB.Recordset")
	Set Rs_give = Server.CreateObject("ADODB.Recordset")
    Set Rs_dct = Server.CreateObject("ADODB.Recordset")
	dbconn.open dbconnect

    dbconn.BeginTrans
    
	sql = "DELETE FROM pay_month_give "& _
	      " WHERE pmg_yymm = '"&pay_month&"' "& _
	      "   AND pmg_id = '1' "& _
	      "   AND pmg_company = '"&pay_company&"'"
	dbconn.execute(sql)
	
	sql = "DELETE FROM pay_month_deduct "&_
	      " WHERE de_yymm = '"&pay_month&"' "&_
	      "   AND de_id = '1' "&_
	      "   AND de_company = '"&pay_company&"'"
	dbconn.execute(sql)
	
  if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "삭제중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "삭제되었습니다...."
  end if

    url = "insa_pay_month_up.asp?ck_sw=y&pay_company=" + pay_company + "&pay_month="+ pay_month
'	if view_condi = "성명" then 
'	           url = "insa_master_modify.asp?ck_sw=y&condi=" + emp_name + "&view_condi="+ view_condi
'		else 
'		       url = "insa_master_modify.asp?ck_sw=y&condi=" + emp_no + "&view_condi="+ view_condi
'	end if
    
    Response.write"<script language=javascript>"
	Response.write"alert('"&end_msg&"');"
'	Response.write"location.replace('insa_pay_month_up.asp');"
	Response.write"location.replace('"&url&"');"
	Response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>
