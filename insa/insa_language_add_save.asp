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
Dim u_type, lang_seq, lang_empno, lang_id, lang_id_type, lang_point
Dim lang_grade, lang_get_date, end_msg, max_seq, rsSeq

u_type = Request.Form("u_type")
lang_seq = Request.Form("lang_seq")
lang_empno = Request.Form("lang_empno")
lang_id = Request.Form("lang_id")
lang_id_type = Request.Form("lang_id_type")
lang_point = Request.Form("lang_point")
lang_grade = Request.Form("lang_grade")
lang_get_date = Request.Form("lang_get_date")

DBConn.BeginTrans

If u_type = "U" Then
	objBuilder.Append "UPDATE emp_language SET "
	objBuilder.Append "	lang_id='"&lang_id&"',lang_id_type='"&lang_id_type&"',lang_point='"&lang_point&"',"
	objBuilder.Append "	lang_grade='"&lang_grade&"',lang_get_date='"&lang_get_date&"',lang_mod_date=now(),lang_mod_user='"&user_name&"'"
	objBuilder.Append "WHERE lang_empno ='"&lang_empno&"' AND lang_seq = '"&lang_seq&"';"
Else
	objBuilder.Append "SELECT MAX(lang_seq) AS 'max_seq' FROM emp_language WHERE lang_empno='"&lang_empno&"';"

	Set rsSeq = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If f_toString(rsSeq("max_seq"), "") = "" Then
		lang_seq = "001"
	Else
		max_seq = "00"&CStr((Int(rsSeq("max_seq")) + 1))
		lang_seq = Right(max_seq,3)
	End If
	rsSeq.Close() : Set rsSeq = Nothing

	objBuilder.Append "INSERT INTO emp_language(lang_empno,lang_seq,lang_id,lang_id_type,lang_point,"
	objBuilder.Append "lang_grade,lang_get_date,lang_reg_date,lang_reg_user)"
	objBuilder.Append "VALUES('"&lang_empno&"','"&lang_seq&"','"&lang_id&"','"&lang_id_type&"','"&lang_point&"',"
	objBuilder.Append "'"&lang_grade&"','"&lang_get_date&"',NOW(),'"&user_name&"');"
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

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	self.opener.location.reload();"
Response.Write "	window.close();"
Response.Write "</script>"
Response.End
%>
