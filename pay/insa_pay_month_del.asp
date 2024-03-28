<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim emp_name, pmg_yymm, view_condi, pmg_id, rs_give
Dim end_msg, page_url

emp_no = Request.Form("in_empno1")
emp_name = Request.Form("in_name1")
pmg_yymm = Request.Form("pmg_yymm1")
view_condi = Request.Form("view_condi1")

pmg_id = "1"

objBuilder.Append "SELECT pmgt.pmg_emp_no, pmdt.de_emp_no "
objBuilder.Append "FROM pay_month_give AS pmgt "
objBuilder.Append "LEFT OUTER JOIN  pay_month_deduct AS pmdt ON pmgt.pmg_emp_no = pmdt.de_emp_no "
objBUilder.Append "	AND pmdt.de_id = '1' AND pmdt.de_yymm = '"&pmg_yymm&"' AND de_company = '"&view_condi&"' "
objBuilder.Append "WHERE pmg_yymm = '"&pmg_yymm&"' AND pmg_emp_no = '"&emp_no&"' AND pmg_id = '1' "
objBuilder.Append "	AND pmg_company = '"&view_condi&"';"

Set rs_give = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

DBConn.BeginTrans

If Not rs_give.EOF Then
	objBuilder.Append "DELETE FROM pay_month_give "
	objBuilder.Append "WHERE pmg_yymm = '"&pmg_yymm&"' AND pmg_emp_no = '"&emp_no&"' AND pmg_id = '1' AND pmg_company = '"&view_condi&"';"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If f_toString(rs_give("de_emp_no"), "") <> "" Then
		objBuilder.Append "DELETE FROM pay_month_deduct "
		objBuilder.Append "WHERE de_yymm = '"&pmg_yymm&"' AND de_emp_no = '"&emp_no&"' AND de_id = '1' AND de_company = '"&view_condi&"';"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	End If

	If Err.number <> 0 then
		DBConn.RollbackTrans
		end_msg = "삭제 중 Error가 발생하였습니다."
	Else
		DBConn.CommitTrans
		end_msg = "정상적으로 삭제되었습니다."
	End If
Else
	end_msg = "삭제할 내역이 없습니다."
End If

DBConn.Close() : Set DBConn = Nothing

page_url = "/pay/insa_pay_month_pay_mg.asp?view_condi="&view_condi&"&pmg_yymm="&pmg_yymm

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
'response.write "location.replace('insa_master_modify.asp');"
Response.Write "	location.replace('"&page_url&"');"
Response.Write "</script>"
Response.End
%>
