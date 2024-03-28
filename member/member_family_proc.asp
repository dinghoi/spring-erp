<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim f_seq, f_rel, f_name, f_birthday, f_birthday_id
Dim f_job, f_live, f_person1, f_person2, f_tel_ddd, f_tel_no1, f_tel_no2
Dim f_support_yn, f_national, f_witak, f_holt, f_holt_date, f_pensioner
Dim f_serius, f_merit, f_disab, f_children, rsMax, end_msg, max_seq

f_seq = Request.Form("f_seq")
f_rel = Request.Form("f_rel")
f_name = Request.Form("f_name")
f_birthday = f_toString(Request.Form("f_birthday"), "")
f_birthday_id = Request.Form("f_birthday_id")
f_job = Request.Form("f_job")
f_live = Request.Form("f_live")
f_person1 = Request.Form("f_person1")
f_person2 = Request.Form("f_person2")
f_tel_ddd = Request.Form("f_tel_ddd")
f_tel_no1 = Request.Form("f_tel_no1")
f_tel_no2 = Request.Form("f_tel_no2")
f_support_yn = Request.Form("f_support_yn")
f_national = Request.Form("f_national")
f_witak = Request.Form("witak_check")
f_holt = Request.Form("holt_check")
f_holt_date = f_toString(Request.Form("f_holt_date"), "")
f_pensioner = Request.Form("pensioner_check")
f_serius = Request.Form("serius_check")
f_merit = Request.Form("merit_check")
f_disab = Request.Form("disab_check")
f_children = Request.Form("children_check")

If f_birthday = "" Then
   f_birthday = "1900-01-01"
End If

If f_holt_date = "" Then
	 f_holt_date = "1900-01-01"
End If

DBConn.BeginTrans

objBuilder.Append "SELECT MAX(f_seq) AS 'max_seq' FROM member_family "
objBuilder.Append "WHERE m_seq='"&m_seq&"';"

Set rsMax = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If f_toString(rsMax("max_seq"), "") = "" Then
	f_seq = "001"
Else
	max_seq = "00"&CStr((Int(rsMax("max_seq")) + 1))
	f_seq = Right(max_seq, 3)
End If
rsMax.Close() : Set rsMax = Nothing

objBuilder.Append "INSERT INTO member_family(m_seq, f_seq, f_rel, f_name, f_birthday, "
objBuilder.Append "f_birthday_id, f_job, f_live, f_support_yn, f_person1, "
objBuilder.Append "f_person2, f_tel_ddd, f_tel_no1, f_tel_no2, f_national, "
objBuilder.Append "f_disab, f_merit, f_serius, f_pensioner, f_witak, "
objBuilder.Append "f_holt, f_holt_date, f_children, f_reg_date)"
objBuilder.Append "VALUES('"&m_seq&"','"&f_seq&"','"&f_rel&"','"&f_name&"','"&f_birthday&"',"
objBuilder.Append "'"&f_birthday_id&"','"&f_job&"','"&f_live&"','"&f_support_yn&"','"&f_person1&"',"
objBuilder.Append "'"&f_person2&"','"&f_tel_ddd&"','"&f_tel_no1&"','"&f_tel_no2&"','"&f_national&"',"
objBuilder.Append "'"&f_disab&"','"&f_merit&"','"&f_serius&"','"&f_pensioner&"','"&f_witak&"',"
objBuilder.Append "'"&f_holt&"','"&f_holt_date&"','"&f_children&"',NOW());"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "등록 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "정상적으로 등록되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	self.opener.location.reload();"
Response.Write "	window.close();"
Response.Write "</script>"
Response.End
%>
