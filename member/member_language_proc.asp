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
Dim lang_seq, lang_id, lang_id_type
Dim lang_point, lang_grade, lang_get_date, rsMax, max_seq, end_msg

lang_id = Request.Form("lang_id")
lang_id_type = Request.Form("lang_id_type")
lang_point = Request.Form("lang_point")
lang_grade = Request.Form("lang_grade")
lang_get_date = Request.Form("lang_get_date")

DBConn.BeginTrans

objBuilder.Append "SELECT MAX(lang_seq) AS 'max_seq' FROM member_language "
objBuilder.Append "WHERE m_seq = '"&m_seq&"';"

Set rsMax = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If f_toString(rsMax("max_seq"), "") = "" Then
	lang_seq = "001"
Else
	max_seq = "00"&CStr((Int(rsMax("max_seq")) + 1))
	lang_seq = Right(max_seq, 3)
End If

rsMax.Close() : Set rsMax = Nothing

objBuilder.Append "INSERT INTO member_language(m_seq,lang_seq,lang_id,lang_id_type,lang_point,"
objBuilder.Append "lang_grade,lang_get_date)VALUES("
objBuilder.Append "'"&m_seq&"','"&lang_seq&"','"&lang_id&"','"&lang_id_type&"','"&lang_point&"',"
objBuilder.Append "'"&lang_grade&"','"&lang_get_date&"');"

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
