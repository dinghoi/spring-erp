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
Dim u_type, qual_seq, qual_empno, qual_type, qual_grade, qual_pass_date
Dim qual_org, qual_no, qual_passport, qual_pay_id, rsMax, max_seq, end_msg

qual_type = request.form("qual_type")
qual_grade = request.form("qual_grade")
qual_pass_date = request.form("qual_pass_date")
qual_org = request.form("qual_org")
qual_no = request.form("qual_no")
qual_passport = request.form("qual_passport")
qual_pay_id = request.form("qual_pay_id")

DBConn.BeginTrans

objBuilder.Append "SELECT MAX(qual_seq) AS 'max_seq' FROM member_qual "
objBuilder.Append "WHERE m_seq = '"&m_seq&"';"

Set rsMax = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If f_toString(rsMax("max_seq"), "") = "" Then
	qual_seq = "001"
Else
	max_seq = "00"&CStr((Int(rsMax("max_seq")) + 1))
	qual_seq = Right(max_seq, 3)
End If

objBuilder.Append "INSERT INTO member_qual(m_seq,qual_seq,qual_type,qual_grade,qual_pass_date, "
objBuilder.Append "qual_org,qual_no,qual_passport,qual_pay_id)VALUES("
objBuilder.Append "'"&m_seq&"','"&qual_seq&"','"&qual_type&"','"&qual_grade&"','"&qual_pass_date&"',"
objBuilder.Append "'"&qual_org&"','"&qual_no&"','"&qual_passport&"','"&qual_pay_id&"');"

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
