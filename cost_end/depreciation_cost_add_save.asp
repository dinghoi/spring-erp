<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

u_type = request.form("u_type")
slip_seq = request.form("slip_seq")
slip_date = request.form("slip_date")
old_date = request.form("old_date")
org_company = request.form("org_company")
account = request.form("account")
account_item = account
sign_no = request.form("sign_no")
pay_method = "현금"
price = int(request.form("price"))
'	vat_yn = request.form("vat_yn")
vat_yn = "N"
customer = ""

company = "공통"
'	emp_no = request.form("emp_no")
pay_yn = "N"
slip_memo = request.form("slip_memo")
end_yn = request.form("end_yn")
cancel_yn = "N"
if vat_yn = "Y" then
	cost = price / 1.1
	cost_vat = cost * 0.1
	cost_vat = round(cost_vat,0)
	cost = price - cost_vat
  else
	cost_vat = 0
	cost = price
end if
mod_id = request.form("mod_id")
mod_user = request.form("mod_user")
mod_date = request.form("mod_date")

if mod_id <> "" then
	mod_yymmdd = datevalue(mod_date)
	mod_hhmm = formatdatetime(mod_date,4)
	mod_date = cstr(mod_yymmdd) + " " + cstr(mod_hhmm)
end if

slip_gubun = "상각비"

set dbconn = server.CreateObject("adodb.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

dbconn.BeginTrans

if	u_type = "U" then
	sql = "delete from general_cost where slip_date ='"&old_date&"' and slip_seq='"&slip_seq&"'"
	dbconn.execute(sql)
end if

sql="select max(slip_seq) as max_seq from general_cost where slip_date='" + slip_date + "'"
set rs=dbconn.execute(sql)

if	isnull(rs("max_seq"))  then
	slip_seq = "001"
  else
	max_seq = "00" + cstr((int(rs("max_seq")) + 1))
	slip_seq = right(max_seq,3)
end if

if isnull(mod_id) or mod_id = "" then
	sql = "insert into general_cost (slip_date,slip_seq,slip_gubun,emp_company,bonbu,saupbu,team,org_name,reside_place,company,account,account_item,pay_method,price,cost,vat_yn,cost_vat,customer,sign_no,emp_name,emp_no,emp_grade,pay_yn,slip_memo,tax_bill_yn,cost_reg,cancel_yn,end_yn,reg_id,reg_user,reg_date) values ('"&slip_date&"','"&slip_seq&"','"&slip_gubun&"','"&org_company&"','','','','"&org_company&"','','"&company&"','"&account&"','"&account_item&"','"&pay_method&"',"&price&","&cost&",'"&vat_yn&"',"&cost_vat&",'"&org_company&"','','"&user_name&"','"&user_id&"','"&user_grade&"','"&pay_yn&"','"&slip_memo&"','N','0','"&cancel_yn&"','"&end_yn&"','"&user_id&"','"&user_name&"',now())"
	dbconn.execute(sql)
  else
	sql = "insert into general_cost (slip_date,slip_seq,slip_gubun,emp_company,bonbu,saupbu,team,org_name,reside_place,company,account,account_item,pay_method,price,cost,vat_yn,cost_vat,customer,sign_no,emp_name,emp_no,emp_grade,pay_yn,slip_memo,tax_bill_yn,cost_reg,cancel_yn,end_yn,reg_id,reg_user,reg_date,mod_id,mod_user,mod_date) values ('"&slip_date&"','"&slip_seq&"','"&slip_gubun&"','"&org_company&"','','','','"&org_company&"','','','"&account&"','"&account_item&"','"&pay_method&"',"&price&","&cost&",'"&vat_yn&"',"&cost_vat&",'"&org_company&"','','"&user_name&"','"&user_id&"','"&user_grade&"','"&pay_yn&"','"&slip_memo&"','N','0','"&cancel_yn&"','"&end_yn&"','"&user_id&"','"&user_name&"',now(),'"&mod_id&"','"&mod_user&"','"&mod_date&"')"
	dbconn.execute(sql)
end If

if Err.number <> 0 then
	dbconn.RollbackTrans
	end_msg = "등록중 Error가 발생하였습니다."
else
	dbconn.CommitTrans
	end_msg = "등록되었습니다."
end if

Response.write "<script type='text/javascript'>"
Response.write "	alert('"&end_msg&"');"
Response.write "	opener.document.frm.submit();"
Response.write "	window.close();"
Response.write "</script>"
Response.End

dbconn.Close() : Set dbconn = Nothing
%>
