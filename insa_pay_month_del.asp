<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

emp_no=Request.form("in_empno1")
emp_name=Request.form("in_name1")
pmg_yymm=Request.form("pmg_yymm1")
view_condi=Request.form("view_condi1")

pmg_id = "1"

'response.write(emp_no)
'response.write(emp_name)
'response.write(pmg_yymm)
'response.write(view_condi)
'response.End

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_give = Server.CreateObject("ADODB.Recordset")
Set rs_duct = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

sql="select * from pay_month_give where pmg_yymm = '"&pmg_yymm&"' and pmg_emp_no = '"&emp_no&"' and pmg_id = '1' and pmg_company = '"&view_condi&"'"
set rs_give=dbconn.execute(sql)
if not rs_give.eof then
   
    sql = " delete from pay_month_give " & _
	            "  where pmg_yymm = '"&pmg_yymm&"' and pmg_emp_no = '"&emp_no&"' and pmg_id = '1' and pmg_company = '"&view_condi&"'"    
	dbconn.execute(sql)
	
	sql="select * from pay_month_deduct where de_yymm = '"&pmg_yymm&"' and de_emp_no = '"&emp_no&"' and de_id = '1' and de_company = '"&view_condi&"'"
    set rs_duct=dbconn.execute(sql)
    if not rs_duct.eof then
           sql = " delete from pay_month_deduct " & _
	             "  where de_yymm = '"&pmg_yymm&"' and de_emp_no = '"&emp_no&"' and de_id = '1' and de_company = '"&view_condi&"'"    
	       dbconn.execute(sql)
	end if
	
   
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "삭제중 Error가 발생하였습니다...."
	  else    
'		dbconn.CommitTrans
		end_msg = "삭제되었습니다...."
	end if

end if
    
	url = "insa_pay_month_pay_mg.asp?ck_sw=y&view_condi=" + view_condi + "&pmg_yymm="+ pmg_yymm
	
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	'response.write"location.replace('insa_master_modify.asp');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
