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
Dim u_type, insu_id, insu_class, insu_id_name, insu_yyyy, from_amt, to_amt
Dim st_amt,emp_rate, com_rate, tot_rate, insu_comment, emp_user, end_msg

u_type = f_Request("u_type")
insu_id = f_Request("insu_id")
insu_class = f_Request("insu_class")
insu_id_name = f_Request("insu_id_name")
insu_yyyy = f_Request("insu_yyyy")
from_amt = Int(f_Request("from_amt"))
to_amt = Int(f_Request("to_amt"))
st_amt = Int(f_Request("st_amt"))
emp_rate = f_Request("emp_rate")
com_rate = f_Request("com_rate")
tot_rate = f_Request("hap_rate")
insu_comment = f_Request("insu_comment")

'hap_rate = emp_rate + com_rate
'start_time = cstr(start_hh) + cstr(start_mm)

emp_user = user_name

DBConn.BeginTrans

If u_type = "U" Then
	objBuilder.Append "UPDATE pay_insurance SET "
	objBuilder.Append "	from_amt='"&from_amt&"',to_amt ='"&to_amt&"',st_amt ='"&st_amt&"', hap_rate='"&tot_rate&"',emp_rate='"&emp_rate&"', "
	objBuilder.Append "	com_rate='"&com_rate&"',insu_comment='"&insu_comment&"',mod_user='"&emp_user&"',mod_date=now() "
	objBuilder.Append "WHERE insu_yyyy = '"&insu_yyyy&"' AND insu_id = '"&insu_id&"' AND insu_class = '"&insu_class&"' "
Else
	objBuilder.Append "INSERT INTO pay_insurance("
	objBuilder.Append "	insu_yyyy, insu_id, insu_class, insu_id_name, from_amt, to_amt, "
	objBuilder.Append "	st_amt, hap_rate, emp_rate, com_rate, insu_comment, reg_user, reg_date "
	objBuilder.Append ")VALUES("
	objBuilder.Append "'"&insu_yyyy&"','"&insu_id&"','"&insu_class&"','"&insu_id_name&"','"&from_amt&"', '"&to_amt&"',"
	objBuilder.Append "'"&st_amt&"','"&tot_rate&"','"&emp_rate&"','"&com_rate&"','"&insu_comment&"','"&emp_user&"', NOW())"
End If

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "저장 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "저장 되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	parent.opener.location.reload();"
Response.Write "	self.close();"
Response.Write "</script>"
Response.End
%>