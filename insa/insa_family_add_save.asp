<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
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
Dim u_type, family_seq, family_empno, family_rel, family_name
Dim family_birthday, family_birthday_id, family_job, family_live
Dim family_person1, family_person2, family_tel_ddd, family_tel_no1
Dim family_tel_no2, family_support_yn, family_national, family_witak
Dim family_holt, family_holt_date, family_pensioner, family_serius
Dim family_merit, family_disab, family_children, rsSeq, max_seq, end_msg

u_type = Request.Form("u_type")
family_seq = Request.Form("family_seq")
family_empno = Request.Form("family_empno")
family_rel = Request.Form("family_rel")
family_name = Request.Form("family_name")
family_birthday = Request.Form("family_birthday")
family_birthday_id = Request.Form("family_birthday_id")
family_job = Request.Form("family_job")
family_live = Request.Form("family_live")
family_person1 = Request.Form("family_person1")
family_person2 = Request.Form("family_person2")
family_tel_ddd = Request.Form("family_tel_ddd")
family_tel_no1 = Request.Form("family_tel_no1")
family_tel_no2 = Request.Form("family_tel_no2")
family_support_yn = Request.Form("family_support_yn")
family_national = Request.Form("family_national")
family_witak = Request.Form("witak_check")
family_holt = Request.Form("holt_check")
family_holt_date = Request.Form("family_holt_date")
family_pensioner = Request.Form("pensioner_check")
family_serius = Request.Form("serius_check")
family_merit = Request.Form("merit_check")
family_disab = Request.Form("disab_check")
family_children = Request.Form("children_check")

If f_toString(family_birthday, "") = "" Then
   family_birthday = "1900-01-01"
End If

If f_toString(family_holt_date, "") = "" Then
	 family_holt_date = "1900-01-01"
End If

DBConn.BeginTrans

If u_type = "U" Then
	objBuilder.Append "UPDATE emp_family SET "
	objBuilder.Append "	family_rel='"&family_rel&"',family_name='"&family_name&"',family_birthday='"&family_birthday&"',"
	objBuilder.Append "	family_birthday_id='"&family_birthday_id&"',family_job='"&family_job&"',family_live='"&family_live&"',"
	objBuilder.Append "	family_support_yn='"&family_support_yn&"',family_person1='"&family_person1&"',family_person2='"&family_person2&"',"
	objBuilder.Append "	family_tel_ddd='"&family_tel_ddd&"',family_tel_no1='"&family_tel_no1&"',family_tel_no2='"&family_tel_no2&"',"
	objBuilder.Append "	family_national='"&family_national&"',family_witak='"&family_witak&"',family_holt='"&family_holt&"',"
	objBuilder.Append "	family_holt_date='"&family_holt_date&"',family_pensioner='"&family_pensioner&"',family_serius='"&family_serius&"',"
	objBuilder.Append "	family_merit='"&family_merit&"',family_disab='"&family_disab&"',family_children='"&family_children&"',"
	objBuilder.Append "	family_mod_date=NOW(), family_mod_user='"&user_name&"' "
	objBuilder.Append "WHERE family_empno ='"&family_empno&"' AND family_seq = '"&family_seq&"';"
Else
	objBuilder.Append "SELECT MAX(family_seq) AS 'max_seq' FROM emp_family "
	objBuilder.Append "WHERE family_empno='"&family_empno&"';"

	Set rsSeq = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If f_toString(rsSeq("max_seq"), "") = "" Then
		family_seq = "001"
	Else
		max_seq = "00"&CStr((Int(rsSeq("max_seq")) + 1))
		family_seq = Right(max_seq, 3)
	End If

	objBuilder.Append "insert into emp_family (family_empno,family_seq,family_rel,family_name,family_birthday,"
	objBuilder.Append "family_birthday_id,family_job,family_live,family_support_yn,family_person1,"
	objBuilder.Append "family_person2,family_tel_ddd,family_tel_no1,family_tel_no2,family_national,"
	objBuilder.Append "family_disab,family_merit,family_serius,family_pensioner,family_witak,"
	objBuilder.Append "family_holt,family_holt_date,family_children,family_reg_date,family_reg_user)"
	objBuilder.Append "VALUES('"&family_empno&"','"&family_seq&"','"&family_rel&"','"&family_name&"','"&family_birthday&"',"
	objBuilder.Append "'"&family_birthday_id&"','"&family_job&"','"&family_live&"','"&family_support_yn&"','"&family_person1&"',"
	objBuilder.Append "'"&family_person2&"','"&family_tel_ddd&"','"&family_tel_no1&"','"&family_tel_no2&"','"&family_national&"',"
	objBuilder.Append "'"&family_disab&"','"&family_merit&"','"&family_serius&"','"&family_pensioner&"','"&family_witak&"',"
	objBuilder.Append "'"&family_holt&"','"&family_holt_date&"','"&family_children&"',now(),'"&user_name&"');"
End If

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

Response.write "<script type='text/javascript'>"
Response.write "	alert('"&end_msg&"');"
Response.write "	self.opener.location.reload();"
Response.write "	window.close();"
Response.write "</script>"
Response.End
%>
