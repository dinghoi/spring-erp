<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
On Error Resume Next
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim approve_no, slip_gubun, company
Dim account, account_item, slip_memo
Dim end_yn, mg_saupbu, rsTax
Dim slip_date, price, cost, cost_vat, customer, customer_no
Dim pay_method, vat_yn, pay_yn
Dim rsEmp, emp_grade, emp_name, rsGe, max_seq, slip_seq, end_msg

approve_no = Request.Form("approve_no")
slip_gubun = Request.Form("slip_gubun")
company = Request.Form("company")
bonbu = Request.Form("bonbu")
saupbu = Request.Form("saupbu")
team = Request.Form("team")
org_name = Request.Form("org_name")
reside_place = Request.Form("reside_place")
account = Request.Form("account")
account_item = Request.Form("account_item")
slip_memo = Request.Form("slip_memo")
end_yn = Request.Form("end_yn")
mg_saupbu = Request.Form("mg_saupbu")
emp_no = Request.Form("emp_no")

DBConn.BeginTrans

'Sql="select * from tax_bill where approve_no = '"&approve_no&"'"
'Set rs = DBConn.Execute(sql)

objBuilder.Append "SELECT bill_date, owner_company, price, cost, cost_vat, "
'objBuilder.Append "	trade_name, trade_no "
objBuilder.Append "	CONVERT(CONVERT(trade_name USING BINARY) USING utf8) AS trade_name, "
objBuilder.Append "	trade_no "
objBuilder.Append "FROM tax_bill WHERE approve_no = '"&approve_no&"'; "

'Response.write objBuilder.ToString()

Set rsTax = DBConn.Execute(objBuilder.Tostring)
objBuilder.Clear()

'아래에서 customer 값을 재 정의함(해당 쿼리 사용의미없음)[허정호_20220210]
'Sql="select * from trade where trade_no = '"&rs("trade_no")&"'"
'Set rs_trade=DbConn.Execute(Sql)
'if rs_trade.eof or rs_trade.bof then
'	customer = rs("trade_name")
' else
'	customer = rs_trade("trade_name")
'end if

slip_date = rsTax("bill_date")
emp_company = rsTax("owner_company")
price = CDbl(rsTax("price"))
cost = CDbl(rsTax("cost"))
cost_vat = CDbl(rsTax("cost_vat"))
customer = CStr(rsTax("trade_name"))
customer_no = rsTax("trade_no")

'Response.write customer

If IsNull(saupbu) Then
	saupbu = ""
End If

If IsNull(reside_place) Then
	reside_place = ""
End If

pay_method = "현금"
vat_yn = "Y"
pay_yn = "N"

rsTax.Close() : Set rsTax = Nothing

'sql="select * from emp_master where emp_no='"&emp_no&"'"
'set rs_emp=dbconn.execute(sql)
objBuilder.Append "SELECT emp_job, emp_name FROM emp_master WHERE emp_no = '"&emp_no&"' "

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

emp_grade = rsEmp("emp_job")
emp_name = rsEmp("emp_name")

rsEmp.Close() : Set rsEmp = Nothing

'sql = "select max(slip_seq) as max_seq from general_cost where slip_date='"&slip_date&"'"
'set rs=dbconn.execute(sql)

objBuilder.Append "SELECT MAX(slip_seq) AS 'max_seq' FROM general_cost WHERE slip_date='"&slip_date&"' "

Set rsGe = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsGe("max_seq")) Then
	slip_seq = "001"
Else
	max_seq = "00" & CStr((Int(rsGe("max_seq")) + 1))
	slip_seq = Right(max_seq, 3)
End If

rsGe.Close() : Set rsGe = Nothing

' 2019.02.02 [박성민 요청] "하장호"의 일반경비 등록시 트랜젹션문제로 ISERT는 되나 UPDATE가 안돼는 문제발생..	트랜잭션 TEST해볼것...
'sql = "insert into general_cost (slip_date,slip_seq,slip_gubun,emp_company,bonbu,saupbu,team,org_name,reside_place,company,account,account_item"&",pay_method,price,cost,vat_yn,cost_vat,customer,customer_no,emp_name,emp_no,emp_grade,pay_yn,slip_memo,tax_bill_yn,cancel_yn,end_yn,reg_id,reg_user,reg_date,approve_no,mg_saupbu) values "&"('"&slip_date&"','"&slip_seq&"','"&slip_gubun&"','"&emp_company&"','"&bonbu&"','"&saupbu&"','"&team&"','"&org_name&"','"&reside_place&"','"&company&"','"&account&"','"&account_item&"','"&pay_method&"',"&price&","&cost&",'"&vat_yn&"',"&cost_vat&",'"&customer&"','"&customer_no&"','"&emp_name&"','"&emp_no&"','"&emp_grade&"','"&pay_yn&"','"&slip_memo&"','Y','N','"&end_yn&"','"&user_id&"','"&user_name&"',now(),'"&approve_no&"','"&mg_saupbu&"')"

objBuilder.Append "INSERT INTO general_cost("
objBuilder.Append "slip_date, slip_seq, slip_gubun, emp_company,bonbu, "
objBuilder.Append "saupbu, team, org_name, reside_place, company, "
objBuilder.Append "account, account_item, pay_method, price, cost, "
objBuilder.Append "vat_yn, cost_vat, customer, customer_no, emp_name, "
objBuilder.Append "emp_no, emp_grade, pay_yn, slip_memo, tax_bill_yn, "
objBuilder.Append "cancel_yn, end_yn, reg_id, reg_user, reg_date, "
objBuilder.Append "approve_no, mg_saupbu "
objBuilder.Append ")VALUES("
objBuilder.Append "'"&slip_date&"','"&slip_seq&"','"&slip_gubun&"','"&emp_company&"','"&bonbu&"', "
objBuilder.Append "'"&saupbu&"','"&team&"','"&org_name&"','"&reside_place&"','"&company&"', "
objBuilder.Append "'"&account&"','"&account_item&"','"&pay_method&"',"&price&","&cost&", "
objBuilder.Append "'"&vat_yn&"',"&cost_vat&",'"&customer&"','"&customer_no&"','"&emp_name&"', "
objBuilder.Append "'"&emp_no&"','"&emp_grade&"','"&pay_yn&"','"&slip_memo&"','Y', "
objBuilder.Append "'N','"&end_yn&"','"&user_id&"','"&user_name&"',NOW(), "
objBuilder.Append "'"&approve_no&"','"&mg_saupbu&"');"

'Response.write objBuilder.Tostring()

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "Update tax_bill set cost_reg_yn='Y',mod_id='"&user_id&"',mod_name='"&user_name&"',mod_date=now() where approve_no = '"&approve_no&"'"
objBuilder.Append "UPDATE tax_bill SET "
objBuilder.Append "	cost_reg_yn='Y', mod_id='"&user_id&"', mod_name='"&user_name&"', mod_date = NOW() "
objBuilder.Append "WHERE approve_no = '"&approve_no&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "등록 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "등록 되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	self.opener.location.reload();"
Response.Write "	window.close();"
Response.Write "</script>"
Response.End
%>
