<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--include virtual="/include/db_create.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
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
Dim title_line
Dim slip_date, slip_seq, slip_gubun, company
Dim account, account_item, slip_memo, mg_saupbu, pl_yn, account_view
Dim sql, end_msg

'	on Error resume next

DBConn.BeginTrans

slip_date = f_Request("slip_date")
slip_seq = f_Request("slip_seq")
slip_gubun = f_Request("slip_gubun")
bonbu = f_Request("bonbu")
saupbu = f_Request("saupbu")
team = f_Request("team")
org_name = f_Request("org_name")
reside_place = f_Request("reside_place")
emp_no = f_Request("emp_no")
company = f_Request("company")
account = f_Request("account")
account_item = f_Request("account_item")
slip_memo = f_Request("slip_memo")
mg_saupbu = f_Request("mg_saupbu")
pl_yn = f_Request("pl_yn")

If slip_gubun = "비용" Then
	account_view = account & "-" & account_item
Else
  	account_view = account_item
End If

title_line = "매입 세금계산서 수정("&account_view&") 등록"

If IsNull(reside_place) Then
	reside_place = ""
End If

Dim rs_emp, emp_grade, emp_name

sql = "SELECT emp_job, emp_name FROM emp_master WHERE emp_no='"&emp_no&"' "
Set rs_emp = DBConn.Execute(sql)

emp_grade = rs_emp("emp_job")
emp_name = rs_emp("emp_name")

rs_emp.Close() : Set rs_emp = Nothing

sql = "Update general_cost set slip_gubun='"&slip_gubun&"',bonbu='"&bonbu&"',saupbu='"&saupbu&"',team='"&team&"',org_name='"&org_name&"',reside_place='"&reside_place&"',company='"&company&"',emp_name='"&emp_name&"',emp_no='"&emp_no&"',emp_grade='"&emp_grade&"',account='"&account&"',account_item='"&account_item&"',slip_memo='"&slip_memo&"',mod_id='"&user_id&"',mod_user='"&user_name&"',mod_date=now(),mg_saupbu = '"&mg_saupbu&"',pl_yn = '"&pl_yn&"' where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"

'DBConn.execute(sql)
%>
<!--#include virtual="/common/log_sales_profit.asp" -->
<%
If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "등록중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "등록되었습니다."
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	self.opener.location.reload();"
Response.Write "	window.close();"
Response.Write "</script>"
Response.End

DBConn.Close() : Set dbconn = Nothing
%>
