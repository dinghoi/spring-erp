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
Dim edu_seq, edu_start_date, edu_end_date
Dim edu_name, edu_office, edu_finish_no, edu_comment, edu_pay
Dim sqlStr, rsMax, max_seq, end_msg

edu_seq = Request.Form("edu_seq")
edu_start_date = Request.Form("edu_start_date")
edu_end_date = Request.Form("edu_end_date")
edu_name = Request.Form("edu_name")
edu_office = Request.Form("edu_office")
edu_finish_no = Request.Form("edu_finish_no")
edu_comment = Request.Form("edu_comment")

'edu_pay = request.form("edu_pay")
edu_pay = 0

DBConn.BeginTrans

sqlStr = "SELECT MAX(edu_seq) AS 'max_seq' FROM member_edu WHERE m_seq = '"&m_seq&"';"
Set rsMax = DBConn.Execute(sqlStr)

If f_toString(rsMax("max_seq"), "") = "" Then
	edu_seq = "001"
Else
	max_seq = "00"&CStr((Int(rsMax("max_seq")) + 1))
	edu_seq = Right(max_seq,3)
End If

objBuilder.Append "INSERT INTO member_edu(m_seq,edu_seq,edu_name,edu_office,edu_finish_no,"
objBuilder.Append "edu_start_date,edu_end_date,edu_pay,edu_comment)VALUES("
objBuilder.Append "'"&m_seq&"','"&edu_seq&"','"&edu_name&"','"&edu_office&"','"&edu_finish_no&"',"
objBuilder.Append "'"&edu_start_date&"','"&edu_end_date&"','"&edu_pay&"','"&edu_comment&"');"

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
