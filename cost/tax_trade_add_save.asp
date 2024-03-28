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
Dim u_type, trade_code, trade_no1, trade_no2, trade_no3
Dim old_trade_no, tradename, trade_id, sales_type, trade_owner
Dim trade_addr, trade_uptae, trade_upjong, trade_tel, trade_fax
Dim group_name, person_name, person_grade, person_tel_no, person_email
Dim person_memo, emp_name, trade_no, trade_name, use_sw

u_type = Request.Form("u_type")
trade_code = Request.Form("trade_code")
trade_no1 = Request.Form("trade_no1")
trade_no2 = Request.Form("trade_no2")
trade_no3 = Request.Form("trade_no3")
old_trade_no = Request.Form("old_trade_no")
tradename = Request.Form("trade_name")
trade_id = Request.Form("trade_id")
sales_type = Request.Form("sales_type")
trade_owner = Request.Form("trade_owner")
trade_addr = Request.Form("trade_addr")
trade_uptae = Request.Form("trade_uptae")
trade_upjong = Request.Form("trade_upjong")
trade_tel = Request.Form("trade_tel")
trade_fax = Request.Form("trade_fax")
group_name = Request.Form("group_name")
person_name = Request.Form("person_name")
person_grade = Request.Form("person_grade")
person_tel_no = Request.Form("person_tel_no")
person_email = Request.Form("person_email")
person_memo = Request.Form("person_memo")
emp_no = Request.Form("emp_no")
emp_name = Request.Form("emp_name")
saupbu = Request.Form("saupbu")

trade_no = CStr(trade_no1) & CStr(trade_no2) & CStr(trade_no3)
trade_name = Replace(tradename,"（주）","(주)")
use_sw = "Y"

DBConn.BeginTrans

Dim sqlStr, rsStr, max_seq, sqlTrade, rsTrade, end_msg

sqlStr = "SELECT MAX(trade_code) AS 'max_seq' FROM trade"
Set rsStr = DBConn.Execute(sqlStr)

If IsNull(rsStr("max_seq")) Then
	trade_code = "00001"
Else
	max_seq = "0000" & CStr((Int(rsStr("max_seq")) + 1))
	trade_code = Right(max_seq, 5)
End If

rsStr.Close() : Set rsStr = Nothing

'sqlTrade = "SELECT trade_no FROM trade WHERE trade_no ='"&trade_no&"' AND (trade_name ='"&trade_name&"' OR trade_full_name ='"&trade_full_name&"')"
sqlTrade = "SELECT trade_no FROM trade WHERE trade_no ='"&trade_no&"' AND trade_name ='"&trade_name&"' "

Set rsTrade = DBConn.Execute(sqlTrade)

If rsTrade.EOF Or rsTrade.BOF Then
	'sql="insert into trade (trade_code,trade_no,trade_name,trade_id,sales_type,trade_owner,trade_addr,trade_uptae,trade_upjong,trade_tel,trade_fax,mg_group,group_name,emp_no,emp_name,saupbu,use_sw,reg_id,reg_date) values ('"&trade_code&"','"&trade_no&"','"&trade_name&"','"&trade_id&"','"&sales_type&"','"&trade_owner&"','"&trade_addr&"','"&trade_uptae&"','"&trade_upjong&"','"&trade_tel&"','"&trade_fax&"','"&mg_group&"','"&group_name&"','"&emp_no&"','"&emp_name&"','"&saupbu&"','"&use_sw&"','"&user_id&"',now())"
	objBuilder.Append "INSERT INTO trade("
	objBuilder.Append "trade_code, trade_no, trade_name, trade_id, sales_type, "
	objBuilder.Append "trade_owner, trade_addr, trade_uptae, trade_upjong, trade_tel, "
	objBuilder.Append "trade_fax, mg_group, group_name, emp_no, emp_name, "
	objBuilder.Append "saupbu, use_sw, reg_id, reg_date"
	objBuilder.Append ")VALUES("
	objBuilder.Append "'"&trade_code&"','"&trade_no&"','"&trade_name&"','"&trade_id&"','"&sales_type&"', "
	objBuilder.Append "'"&trade_owner&"','"&trade_addr&"','"&trade_uptae&"','"&trade_upjong&"','"&trade_tel&"', "
	objBuilder.Append "'"&trade_fax&"','"&mg_group&"','"&group_name&"','"&emp_no&"','"&emp_name&"', "
	objBuilder.Append "'"&saupbu&"','"&use_sw&"','"&user_id&"',now())"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If (person_name <> "" Or IsNull(person_name)) And (person_email <> "" Or IsNull(person_name)) Then
		'sql="insert into trade_person (trade_code,person_name,person_grade,person_tel_no,person_email,person_memo,reg_id,reg_name,reg_date) values ('"&trade_code&"','"&person_name&"','"&person_grade&"','"&person_tel_no&"','"&person_email&"','"&person_memo&"','"&user_id&"','"&user_name&"',now())"
		objBuilder.Append "INSERT INTO trade_person("
		objBuilder.Append "trade_code, person_name, person_grade, person_tel_no, person_email,"
		objBuilder.Append "person_memo, reg_id, reg_name, reg_date"
		objBuilder.Append ")VALUES("
		objBuilder.Append "'"&trade_code&"', '"&person_name&"', '"&person_grade&"', '"&person_tel_no&"', '"&person_email&"',"
		objBuilder.Append "'"&person_memo&"', '"&user_id&"', '"&user_name&"', NOW())"

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	End If
Else
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('이미 등록된 거래처입니다');"
	Response.Write "	history.back();"
	Response.Write "</script>"
	Response.End
End If

rsTrade.Close() : Set rsTrade = Nothing

If Err.number <> 0 Then
	DBConn.RollbackTrans
	'end_msg = sms_msg & "처리중 Error가 발생하였습니다...."
	end_msg = "처리중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	'end_msg = sms_msg & "처리 되었습니다...."
	end_msg = "처리 되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	opener.document.frm.submit();"
Response.Write "	self.close() ;"
Response.Write "</script>"
Response.End
%>
