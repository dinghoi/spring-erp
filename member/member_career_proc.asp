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
Dim c_join_date, c_end_date,c_office, c_dept, c_position, c_task
Dim rsMax, max_seq, c_seq, end_msg

c_join_date = request.form("c_join_date")
c_end_date = request.form("c_end_date")
c_office = request.form("c_office")
c_dept = request.form("c_dept")
c_position = request.form("c_position")
c_task = request.form("c_task")

DBConn.BeginTrans

objBuilder.Append "SELECT MAX(c_seq) AS 'max_seq' FROM member_career "
objBuilder.Append "WHERE m_seq = '"&m_seq&"';"

Set rsMax = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If f_toString(rsMax("max_seq"), "") = ""  Then
	c_seq = "001"
Else
	max_seq = "00"&CStr((Int(rsMax("max_seq")) + 1))
	c_seq = Right(max_seq, 3)
End If

objBuilder.Append "INSERT INTO member_career(m_seq, c_seq, c_join_date, c_end_date, c_office, c_dept, c_position, c_task)VALUES("
objBuilder.Append "'"&m_seq&"', '"&c_seq&"','"&c_join_date&"','"&c_end_date&"','"&c_office&"',"
objBuilder.Append "'"&c_dept&"','"&c_position&"','"&c_task&"');"

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
